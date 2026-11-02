/**
 * Project AI (pilot: business plan) — shell main-process service.
 *
 * Flow: the user picks a saved project folder and types a short brief.
 * The service generates a plan document (.docx), a budget spreadsheet
 * (.xlsx) and a pitch deck (.pptx) into that folder using the same
 * bundled-CLI paths the headless MCP tools use (createFileViaCli /
 * createPptxViaCli), registers the files with the existing project
 * store, and opens them as tabs. It then watches the folder (debounced)
 * and re-runs a cross-file review whenever the files change.
 *
 * Safety: the deterministic review in `../shared/project-ai` is the only
 * source of auto-applied changes, and only for typo fixes and number
 * syncs where the spreadsheet is the unambiguous source of truth. The
 * language model only ever produces approval-queue suggestions, and any
 * model or file failure degrades to the template review rather than
 * failing the build. No paid service, no key requirement, no telemetry:
 * the model call reuses the user's own configured provider (local
 * Ollama by default) and is skipped entirely when no model is set.
 */

import { createHash } from 'node:crypto'
import { existsSync, readFileSync, watch, type FSWatcher } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, join } from 'node:path'
import JSZip from 'jszip'
import {
  activeProvider,
  chatForProvider,
  defaultAiSettings,
  resolveAiSettings,
} from '@genoffice/ai-provider'
import { ProjectStore } from '@genoffice/project-store'
import { atomicWriteFile } from './atomic-write'
import type { CliRunner } from './mcp/cli-runner'
import { createFileViaCli, createPptxViaCli, readDocxTextViaCli } from './mcp/headless-cli'
import { outlineToOps, parsePptxOutline } from './mcp/pptx-outline'
import {
  formatMoney,
  PROJECT_AI_DECK,
  PROJECT_AI_DOC,
  PROJECT_AI_SHEET,
  buildBusinessPlanContent,
  businessNameFromBrief,
  businessPlanBudgetRows,
  businessPlanDeckOutline,
  businessPlanMarkdown,
  budgetTotal,
  projectAiFileName,
  reviewProjectContent,
  type BudgetRow,
  type BusinessPlanDraft,
  type ProjectAiChangeLogEntry,
  type ProjectAiFileKind,
  type ProjectAiState,
  type ProjectAiSuggestion,
} from '../shared/project-ai'

/** Debounce for folder-change reviews: never on every keystroke/save. */
export const PROJECT_AI_REVIEW_DEBOUNCE_MS = 1500

export interface ProjectAiServiceDeps {
  cli: CliRunner
  userDataPath: string
  /** open a file path as a tab (routeDocumentPath in the shell main) */
  openFile: (path: string) => void
  /** model draft override (tests inject a fake; production uses Ollama/BYOK) */
  draftWithAi?: (brief: string) => Promise<BusinessPlanDraft | null>
  /** model review override (tests inject a fake) */
  suggestWithAi?: (input: {
    docText: string
    deckText: string
    budgetRows: BudgetRow[]
  }) => Promise<ProjectAiSuggestion[]>
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function validateDraft(value: unknown, brief: string): BusinessPlanDraft | null {
  if (!isRecord(value)) return null
  const sectionsRaw = value.sections
  const budgetRaw = value.budget
  if (!Array.isArray(sectionsRaw) || !Array.isArray(budgetRaw)) return null
  const sections: BusinessPlanDraft['sections'] = []
  for (const entry of sectionsRaw.slice(0, 8)) {
    if (!isRecord(entry)) return null
    const heading = typeof entry.heading === 'string' ? entry.heading.trim() : ''
    const body = typeof entry.body === 'string' ? entry.body.trim() : ''
    if (!heading || !body) return null
    sections.push({ heading, body })
  }
  const budget: BudgetRow[] = []
  for (const entry of budgetRaw.slice(0, 20)) {
    if (!isRecord(entry)) return null
    const label = typeof entry.label === 'string' ? entry.label.trim() : ''
    const amount = typeof entry.amount === 'number' ? entry.amount : Number(entry.amount)
    if (!label || !Number.isFinite(amount) || amount <= 0) return null
    budget.push({ label, amount })
  }
  if (sections.length === 0 || budget.length === 0) return null
  const businessName =
    typeof value.businessName === 'string' && value.businessName.trim()
      ? value.businessName.trim()
      : businessNameFromBrief(brief)
  return { businessName, sections, budget }
}

function extractJson(text: string): unknown | null {
  const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(text)
  const candidate = (fenced ? fenced[1] : text).trim()
  try {
    return JSON.parse(candidate) as unknown
  } catch {
    const start = candidate.indexOf('{')
    const end = candidate.lastIndexOf('}')
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(candidate.slice(start, end + 1)) as unknown
      } catch {
        return null
      }
    }
    return null
  }
}

/** Resolve the user's configured provider (local Ollama by default).
 *  Returns null when no usable model is configured — callers fall back. */
function resolveModelTarget(userDataPath: string): {
  provider: Parameters<typeof chatForProvider>[0]
  config: Parameters<typeof chatForProvider>[1]
} | null {
  let stored: unknown
  try {
    stored = JSON.parse(readFileSync(join(userDataPath, 'ai-settings.json'), 'utf8')) as unknown
  } catch {
    stored = {}
  }
  const settings = resolveAiSettings(stored, defaultAiSettings())
  const provider = activeProvider(settings)
  // Never route a project request to Genspark: activeProvider already
  // migrates it to Custom, and this is a second guard.
  if (provider === 'genspark') return null
  const config = settings.providers?.[provider]
  if (!config) return null
  if (provider === 'custom' && !config.model?.trim()) return null
  return { provider, config }
}

async function callModel(
  userDataPath: string,
  system: string,
  user: string,
  timeoutMs = 30_000,
): Promise<string | null> {
  const target = resolveModelTarget(userDataPath)
  if (!target) return null
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    const result = await chatForProvider(
      target.provider,
      target.config,
      system,
      user,
      controller.signal,
    )
    if (!result.ok) return null
    return typeof result.text === 'string' && result.text.trim() ? result.text : null
  } catch {
    return null
  } finally {
    clearTimeout(timer)
  }
}

const DRAFT_SYSTEM =
  'You draft small business plans. Reply with JSON only, no prose: ' +
  '{"businessName": "...", "sections": [{"heading": "...", "body": "..."}], ' +
  '"budget": [{"label": "...", "amount": 0}]}. Keep sections to at most 5 and ' +
  'budget rows to at most 10. Amounts are US dollars, numbers only.'

const REVIEW_SYSTEM =
  'You review a business plan project: a plan document, a pitch deck, and a ' +
  'budget spreadsheet (the source of truth for figures). Reply with JSON only, ' +
  'no prose: an array of {"title": "...", "detail": "...", "file": "doc" | "deck", ' +
  '"before": "exact text in that file", "after": "replacement text"}. Suggest only ' +
  'wording, structure, or new-content improvements — never typo fixes and never ' +
  'number corrections, which are handled elsewhere. If nothing is worth ' +
  'changing, reply with [].'

function validateModelSuggestions(value: unknown, startIndex: number): ProjectAiSuggestion[] {
  if (!Array.isArray(value)) return []
  const out: ProjectAiSuggestion[] = []
  let index = startIndex
  for (const entry of value.slice(0, 20)) {
    if (!isRecord(entry)) continue
    const file = entry.file === 'doc' || entry.file === 'deck' ? entry.file : null
    const title = typeof entry.title === 'string' ? entry.title.trim() : ''
    const detail = typeof entry.detail === 'string' ? entry.detail.trim() : ''
    const before = typeof entry.before === 'string' ? entry.before : ''
    const after = typeof entry.after === 'string' ? entry.after : ''
    if (!file || !title || !before || !after || before === after) continue
    index += 1
    // Model suggestions are ALWAYS queue items, even if they look like a
    // typo or a number fix: only deterministic findings may auto-apply.
    out.push({
      id: `other-${file}-${index}`,
      kind: 'other',
      file,
      fileName: projectAiFileName(file),
      title,
      detail,
      before,
      after,
      autoSafe: false,
    })
  }
  return out
}

/** Collect every `text` string in a slides-read JSON tree. */
export function collectDeckText(detail: unknown): string {
  const parts: string[] = []
  const walk = (value: unknown): void => {
    if (typeof value === 'string') return
    if (Array.isArray(value)) {
      for (const entry of value) walk(entry)
      return
    }
    if (isRecord(value)) {
      if (typeof value.text === 'string' && value.text) parts.push(value.text)
      for (const entry of Object.values(value)) walk(entry)
    }
  }
  walk(detail)
  return parts.join('\n')
}

/** Budget rows from a `sheet read` rows matrix: [label, amount, ...]. */
export function budgetRowsFromSheetRows(rows: unknown): BudgetRow[] {
  if (!Array.isArray(rows)) return []
  const out: BudgetRow[] = []
  for (const row of rows) {
    if (!Array.isArray(row) || row.length < 2) continue
    const label = typeof row[0] === 'string' ? row[0].trim() : ''
    const amount = typeof row[1] === 'number' ? row[1] : Number(row[1])
    if (!label || label.toLowerCase() === 'category') continue
    if (!Number.isFinite(amount)) continue
    out.push({ label, amount })
  }
  return out
}

/**
 * Replace text inside XML text nodes (`<w:t>` / `<a:t>`) only.
 * - typo fixes replace every occurrence (same word, same fix).
 * - number syncs and queue accepts require exactly one occurrence, so a
 *   figure that appears twice for different reasons is never silently
 *   rewritten in both places.
 * Returns the new XML, or null when the occurrence rule is not met.
 */
export function replaceInTextNodes(
  xml: string,
  before: string,
  after: string,
  replaceAll: boolean,
): string | null {
  if (!before) return null
  const escapedAfter = after.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  let occurrences = 0
  const replaced = xml.replace(
    /(<(?:w:t|a:t)(?:\s[^>]*)?>)([^<]*)(<\/(?:w:t|a:t)>)/g,
    (whole, open: string, text: string, close: string) => {
      if (!text.includes(before)) return whole
      const count = text.split(before).length - 1
      occurrences += count
      if (!replaceAll && occurrences > 1) return whole
      return `${open}${text.split(before).join(escapedAfter)}${close}`
    },
  )
  if (occurrences === 0) return null
  if (!replaceAll && occurrences !== 1) return null
  return replaced
}

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex')
}

export class ProjectAiService {
  private readonly deps: ProjectAiServiceDeps
  private state: ProjectAiState = {
    folder: null,
    files: [],
    status: 'idle',
    statusMessage: '',
    applied: [],
    queue: [],
    lastReviewAt: null,
    lastError: null,
  }
  private watcher: FSWatcher | null = null
  private reviewTimer: ReturnType<typeof setTimeout> | null = null
  private reviewing = false
  /** content hashes at the last review, so an unchanged folder never re-reviews */
  private readonly reviewedHashes = new Map<string, string>()

  constructor(deps: ProjectAiServiceDeps) {
    this.deps = deps
  }

  getState(): ProjectAiState {
    return {
      ...this.state,
      files: [...this.state.files],
      applied: [...this.state.applied],
      queue: this.state.queue.map((s) => ({ ...s })),
    }
  }

  /** Build the three project files, register + open them, watch, review. */
  async buildProject(folder: string, brief: string): Promise<{ files: string[] }> {
    const trimmed = brief.trim()
    if (!trimmed) throw new Error('A brief is required to build a business plan.')
    this.stopWatching()
    this.state = {
      folder,
      files: [],
      status: 'building',
      statusMessage: 'Drafting the business plan…',
      applied: [],
      queue: [],
      lastReviewAt: null,
      lastError: null,
    }
    const draft = (await this.draftPlan(trimmed)) ?? buildBusinessPlanContent(trimmed)
    const docPath = join(folder, PROJECT_AI_DOC)
    const sheetPath = join(folder, PROJECT_AI_SHEET)
    const deckPath = join(folder, PROJECT_AI_DECK)
    await createFileViaCli(this.deps.cli, {
      type: 'docx',
      input: { name: 'content.md', content: businessPlanMarkdown(draft) },
      out: docPath,
      overwrite: true,
    })
    await createFileViaCli(this.deps.cli, {
      type: 'xlsx',
      input: {
        name: 'table.json',
        content: JSON.stringify({
          sheets: [{ name: 'Budget', rows: businessPlanBudgetRows(draft) }],
        }),
      },
      out: sheetPath,
      overwrite: true,
    })
    const outline = businessPlanDeckOutline(draft)
    await createPptxViaCli(this.deps.cli, {
      ops: outlineToOps(parsePptxOutline('markdown', outline)),
      out: deckPath,
      overwrite: true,
    })
    const files = [docPath, sheetPath, deckPath]
    // Best effort: the project store is chat history, not the build path —
    // a store failure must never fail a build the user just watched.
    try {
      const store = new ProjectStore(this.deps.userDataPath)
      const project = store.createProject(draft.businessName)
      for (const file of files) store.moveFileToProject(file, project.id)
    } catch {
      // ignore: the folder on disk is the project of record
    }
    this.state.files = files
    this.state.status = 'ready'
    this.state.statusMessage = `Built ${draft.businessName}: plan, budget (${formatMoney(budgetTotal(draft.budget))} total) and pitch deck.`
    for (const file of files) {
      try {
        this.deps.openFile(file)
      } catch {
        // opening a tab must not fail the build either; files are on disk
      }
    }
    this.startWatching(folder)
    await this.runReview()
    return { files }
  }

  private async draftPlan(brief: string): Promise<BusinessPlanDraft | null> {
    if (this.deps.draftWithAi) {
      try {
        return await this.deps.draftWithAi(brief)
      } catch {
        return null
      }
    }
    const text = await callModel(
      this.deps.userDataPath,
      DRAFT_SYSTEM,
      `Business idea / brief:\n${brief}`,
    )
    if (!text) return null
    return validateDraft(extractJson(text), brief)
  }

  startWatching(folder: string): void {
    this.stopWatching()
    try {
      this.watcher = watch(folder, { recursive: true }, () => this.scheduleReview())
    } catch {
      // watching is best effort: the panel can still review on demand
      this.watcher = null
    }
  }

  stopWatching(): void {
    if (this.reviewTimer) {
      clearTimeout(this.reviewTimer)
      this.reviewTimer = null
    }
    if (this.watcher) {
      this.watcher.close()
      this.watcher = null
    }
  }

  /** Debounced folder-change review (never on every keystroke or save). */
  scheduleReview(): void {
    if (this.reviewTimer) clearTimeout(this.reviewTimer)
    this.reviewTimer = setTimeout(() => {
      void this.runReview()
    }, PROJECT_AI_REVIEW_DEBOUNCE_MS)
  }

  /** Read the three files, review them, auto-apply SAFE, queue the rest. */
  async runReview(): Promise<ProjectAiState> {
    const folder = this.state.folder
    if (!folder || this.reviewing) return this.getState()
    this.reviewing = true
    try {
      const docPath = join(folder, PROJECT_AI_DOC)
      const sheetPath = join(folder, PROJECT_AI_SHEET)
      const deckPath = join(folder, PROJECT_AI_DECK)
      if (!existsSync(docPath) || !existsSync(sheetPath) || !existsSync(deckPath)) {
        this.state.lastError = 'Project files are missing from the folder.'
        return this.getState()
      }
      const hashes = new Map<string, string>()
      for (const file of [docPath, sheetPath, deckPath]) {
        try {
          hashes.set(file, sha256(await readFile(file)))
        } catch {
          hashes.set(file, '')
        }
      }
      const unchanged = [docPath, sheetPath, deckPath].every(
        (file) => this.reviewedHashes.get(file) === hashes.get(file),
      )
      if (unchanged && this.state.lastReviewAt) return this.getState()
      for (const [file, hash] of hashes) this.reviewedHashes.set(file, hash)

      const docText = await readDocxTextViaCli(this.deps.cli, docPath)
      const sheetOutcome = await this.deps.cli.run(['sheet', 'read', sheetPath])
      if (!sheetOutcome.ok || !sheetOutcome.json || sheetOutcome.json.status !== 'ok') {
        throw new Error('Could not read the budget spreadsheet.')
      }
      const budgetRows = budgetRowsFromSheetRows(
        (sheetOutcome.json.detail as { rows?: unknown } | undefined)?.rows,
      )
      const deckOutcome = await this.deps.cli.run(['slides', 'read', deckPath, '--full'])
      if (!deckOutcome.ok || !deckOutcome.json || deckOutcome.json.status !== 'ok') {
        throw new Error('Could not read the pitch deck.')
      }
      const deckText = collectDeckText(deckOutcome.json.detail)

      const { safe, queued } = reviewProjectContent({ docText, deckText, budgetRows })
      const appliedNow: ProjectAiChangeLogEntry[] = []
      const failedSafe: ProjectAiSuggestion[] = []
      for (const suggestion of safe) {
        const ok = await this.applyToDisk(suggestion)
        if (ok) {
          appliedNow.push({
            at: new Date().toISOString(),
            suggestionId: suggestion.id,
            kind: suggestion.kind,
            fileName: suggestion.fileName,
            before: suggestion.before,
            after: suggestion.after,
            via: 'auto',
          })
        } else {
          // could not apply exactly (split text runs, repeated figure):
          // never force it — it becomes a queue item instead.
          failedSafe.push({
            ...suggestion,
            autoSafe: false,
            ambiguity: 'could not be applied exactly, so it needs your approval',
          })
        }
      }
      let modelSuggestions: ProjectAiSuggestion[] = []
      try {
        modelSuggestions = this.deps.suggestWithAi
          ? await this.deps.suggestWithAi({ docText, deckText, budgetRows })
          : await this.suggestWithModel(docText, deckText, budgetRows)
      } catch {
        modelSuggestions = []
      }
      const seen = new Set(this.state.queue.map((s) => s.id))
      const nextQueue = [...this.state.queue]
      for (const suggestion of [...failedSafe, ...queued, ...modelSuggestions]) {
        if (!seen.has(suggestion.id)) {
          seen.add(suggestion.id)
          nextQueue.push(suggestion)
        }
      }
      this.state.queue = nextQueue
      this.state.applied = [...this.state.applied, ...appliedNow]
      this.state.lastReviewAt = new Date().toISOString()
      this.state.lastError = null
      return this.getState()
    } catch (error) {
      this.state.lastError = error instanceof Error ? error.message : String(error)
      return this.getState()
    } finally {
      this.reviewing = false
    }
  }

  private async suggestWithModel(
    docText: string,
    deckText: string,
    budgetRows: BudgetRow[],
  ): Promise<ProjectAiSuggestion[]> {
    const excerpt = (text: string): string =>
      text.length > 4000 ? `${text.slice(0, 4000)}…` : text
    const text = await callModel(
      this.deps.userDataPath,
      REVIEW_SYSTEM,
      [
        `Budget (source of truth): ${JSON.stringify(budgetRows)}`,
        `Plan document:\n${excerpt(docText)}`,
        `Pitch deck:\n${excerpt(deckText)}`,
      ].join('\n\n'),
    )
    if (!text) return []
    return validateModelSuggestions(extractJson(text), 0)
  }

  /** Accept a queued suggestion: apply it to disk and log it. */
  async acceptSuggestion(id: string): Promise<boolean> {
    const suggestion = this.state.queue.find((s) => s.id === id)
    if (!suggestion) return false
    if (!suggestion.before) return false
    const ok = await this.applyToDisk(suggestion)
    if (!ok) {
      this.state.lastError = `Could not apply “${suggestion.title}” exactly; the text may have changed.`
      return false
    }
    this.state.queue = this.state.queue.filter((s) => s.id !== id)
    this.state.applied.push({
      at: new Date().toISOString(),
      suggestionId: suggestion.id,
      kind: suggestion.kind,
      fileName: suggestion.fileName,
      before: suggestion.before,
      after: suggestion.after,
      via: 'accepted',
    })
    return true
  }

  rejectSuggestion(id: string): void {
    this.state.queue = this.state.queue.filter((s) => s.id !== id)
  }

  /**
   * Apply one suggestion to the on-disk file (never to the spreadsheet).
   * Typo fixes replace every occurrence; everything else requires exactly
   * one text-node occurrence. Returns false when the rule is not met —
   * the caller queues the suggestion instead of forcing it.
   */
  async applyToDisk(suggestion: ProjectAiSuggestion): Promise<boolean> {
    // the spreadsheet is the source of truth: it is read, never edited here
    if (suggestion.file === 'sheet') return false
    const folder = this.state.folder
    if (!folder) return false
    const path = join(folder, suggestion.fileName)
    if (!existsSync(path)) return false
    try {
      const bytes = await readFile(path)
      const zip = await JSZip.loadAsync(bytes)
      const targets =
        suggestion.file === 'doc'
          ? ['word/document.xml']
          : Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      let changed = false
      for (const name of targets) {
        const entry = zip.files[name]
        if (!entry) continue
        const xml = await entry.async('string')
        const next = replaceInTextNodes(
          xml,
          suggestion.before,
          suggestion.after,
          suggestion.kind === 'typo',
        )
        if (next !== null && next !== xml) {
          zip.file(name, next)
          changed = true
          if (suggestion.kind !== 'typo') break
        }
      }
      if (!changed) return false
      const out = await zip.generateAsync({ type: 'nodebuffer' })
      await atomicWriteFile(path, out)
      this.reviewedHashes.delete(path)
      return true
    } catch {
      return false
    }
  }
}

export function projectAiDocPath(folder: string): string {
  return join(folder, PROJECT_AI_DOC)
}

export function projectAiSheetPath(folder: string): string {
  return join(folder, PROJECT_AI_SHEET)
}

export function projectAiDeckPath(folder: string): string {
  return join(folder, PROJECT_AI_DECK)
}

export function projectAiFileKindForName(name: string): ProjectAiFileKind | null {
  if (name === PROJECT_AI_DOC) return 'doc'
  if (name === PROJECT_AI_SHEET) return 'sheet'
  if (name === PROJECT_AI_DECK) return 'deck'
  return null
}

export function isProjectAiFileName(name: string): boolean {
  return projectAiFileKindForName(basename(name)) !== null
}
