/**
 * Project AI — shell main-process service.
 *
 * One flow serves every project type (business plan, grant application,
 * event plan, marketing plan): the per-type definition in
 * `../shared/project-ai` supplies the file names, document sections,
 * budget shape, deck slides and prompt framing, and this service runs
 * the same build / watch / review steps for each of them.
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
  PROJECT_AI_PROJECT_TYPES,
  formatMoney,
  buildProjectContent,
  budgetTotal,
  normalizeProjectDraft,
  normalizeProjectType,
  projectAiFileName,
  projectAiTypeDefinition,
  projectDraftBudgetRows,
  projectDraftDeckOutline,
  projectDraftMarkdown,
  reviewProjectContent,
  validateModelSuggestions,
  validateProjectDraft,
  type BudgetRow,
  type ProjectAiChangeLogEntry,
  type ProjectAiFileKind,
  type ProjectAiProjectType,
  type ProjectAiState,
  type ProjectAiSuggestion,
  type ProjectDraft,
} from '../shared/project-ai'

/** Debounce for folder-change reviews: never on every keystroke/save. */
export const PROJECT_AI_REVIEW_DEBOUNCE_MS = 1500

export interface ProjectAiServiceDeps {
  cli: CliRunner
  userDataPath: string
  /** open a file path as a tab (routeDocumentPath in the shell main) */
  openFile: (path: string) => void
  /** model draft override (tests inject a fake; production uses Ollama/BYOK) */
  draftWithAi?: (brief: string, projectType: ProjectAiProjectType) => Promise<ProjectDraft | null>
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

/** Draft validation is the shared, bounded validateProjectDraft. */

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
    return typeof result.content === 'string' && result.content.trim() ? result.content : null
  } catch {
    return null
  } finally {
    clearTimeout(timer)
  }
}

/** Model draft prompt for a project type, framed by its definition. */
export function draftSystemFor(projectType: unknown): string {
  const definition = projectAiTypeDefinition(projectType)
  return (
    `You draft ${definition.draftSubject}. Reply with JSON only, no prose: ` +
    '{"title": "...", "sections": [{"heading": "...", "body": "..."}], ' +
    '"budget": [{"label": "...", "amount": 0}]}. Keep sections to at most 5 and ' +
    'budget rows to at most 10. Amounts are US dollars, numbers only.'
  )
}

/** Model review prompt for a project type, framed by its definition. */
export function reviewSystemFor(projectType: unknown): string {
  const definition = projectAiTypeDefinition(projectType)
  return (
    `You review a ${definition.reviewSubject}: a project document, a deck, and a ` +
    'budget spreadsheet (the source of truth for figures). Reply with JSON only, ' +
    'no prose: an array of {"title": "...", "detail": "...", "file": "doc" | "deck", ' +
    '"before": "exact text in that file", "after": "replacement text"}. Suggest only ' +
    'wording, structure, or new-content improvements — never typo fixes and never ' +
    'number corrections, which are handled elsewhere. If nothing is worth ' +
    'changing, reply with [].'
  )
}

/** Model suggestions are validated by the shared validateModelSuggestions:
 *  they are always queue items (kind other, autoSafe false), for every type. */

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
    projectType: 'business-plan',
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

  /** Build the three project files, register + open them, watch, review.
   *  `projectType` may be any stored / incoming value: a missing or
   *  unknown type builds a business plan, so old callers keep working. */
  async buildProject(
    folder: string,
    brief: string,
    projectTypeInput?: unknown,
  ): Promise<{ files: string[] }> {
    const projectType = normalizeProjectType(projectTypeInput)
    const definition = projectAiTypeDefinition(projectType)
    const trimmed = brief.trim()
    if (!trimmed) throw new Error('A brief is required to build a project.')
    this.stopWatching()
    this.state = {
      folder,
      files: [],
      projectType,
      status: 'building',
      statusMessage: `Drafting the ${definition.docTitle.toLowerCase()}…`,
      applied: [],
      queue: [],
      lastReviewAt: null,
      lastError: null,
    }
    const draft =
      (await this.draftPlan(trimmed, projectType)) ?? buildProjectContent(projectType, trimmed)
    const docPath = join(folder, projectAiFileName('doc', projectType))
    const sheetPath = join(folder, projectAiFileName('sheet', projectType))
    const deckPath = join(folder, projectAiFileName('deck', projectType))
    await createFileViaCli(this.deps.cli, {
      type: 'docx',
      input: { name: 'content.md', content: projectDraftMarkdown(draft) },
      out: docPath,
      overwrite: true,
    })
    await createFileViaCli(this.deps.cli, {
      type: 'xlsx',
      input: {
        name: 'table.json',
        content: JSON.stringify({
          sheets: [{ name: definition.sheetName, rows: projectDraftBudgetRows(draft) }],
        }),
      },
      out: sheetPath,
      overwrite: true,
    })
    const outline = projectDraftDeckOutline(draft)
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
    this.state.statusMessage = `Built ${draft.businessName}: ${definition.buildSummary
      .split('{total}')
      .join(formatMoney(budgetTotal(draft.budget)))}.`
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

  private async draftPlan(
    brief: string,
    projectType: ProjectAiProjectType,
  ): Promise<ProjectDraft | null> {
    if (this.deps.draftWithAi) {
      try {
        const draft = await this.deps.draftWithAi(brief, projectType)
        // An injected draft may predate project types (no type / title):
        // carry it into the shape for the type actually being built.
        return draft ? normalizeProjectDraft(draft, projectType) : null
      } catch {
        return null
      }
    }
    const definition = projectAiTypeDefinition(projectType)
    const text = await callModel(
      this.deps.userDataPath,
      draftSystemFor(projectType),
      `${definition.briefPromptLabel}\n${brief}`,
    )
    if (!text) return null
    return validateProjectDraft(extractJson(text), brief, projectType)
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
      const projectType = normalizeProjectType(this.state.projectType)
      const docPath = join(folder, projectAiFileName('doc', projectType))
      const sheetPath = join(folder, projectAiFileName('sheet', projectType))
      const deckPath = join(folder, projectAiFileName('deck', projectType))
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

      const { safe, queued } = reviewProjectContent({ docText, deckText, budgetRows, projectType })
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
        const rawModelSuggestions = this.deps.suggestWithAi
          ? await this.deps.suggestWithAi({ docText, deckText, budgetRows })
          : await this.suggestWithModel(docText, deckText, budgetRows)
        // Belt and braces, for every project type: model output is always
        // a queue item — kind 'other', never autoSafe, and addressed to
        // this project's files — whatever the model (or a test fake
        // standing in for it) returned.
        modelSuggestions = rawModelSuggestions.map((suggestion) => ({
          ...suggestion,
          kind: 'other' as const,
          autoSafe: false,
          fileName: projectAiFileName(suggestion.file, projectType),
        }))
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
    const projectType = normalizeProjectType(this.state.projectType)
    const excerpt = (text: string): string =>
      text.length > 4000 ? `${text.slice(0, 4000)}…` : text
    const text = await callModel(
      this.deps.userDataPath,
      reviewSystemFor(projectType),
      [
        `Budget (source of truth): ${JSON.stringify(budgetRows)}`,
        `Project document:\n${excerpt(docText)}`,
        `Deck:\n${excerpt(deckText)}`,
      ].join('\n\n'),
    )
    if (!text) return []
    return validateModelSuggestions(extractJson(text), 0, projectType)
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

export function projectAiDocPath(folder: string, projectType?: unknown): string {
  return join(folder, projectAiFileName('doc', normalizeProjectType(projectType)))
}

export function projectAiSheetPath(folder: string, projectType?: unknown): string {
  return join(folder, projectAiFileName('sheet', normalizeProjectType(projectType)))
}

export function projectAiDeckPath(folder: string, projectType?: unknown): string {
  return join(folder, projectAiFileName('deck', normalizeProjectType(projectType)))
}

/** Which project type and file kind a file name belongs to, across all
 *  project types; null when the name is not a Project AI file. */
export function projectAiFileInfoForName(
  name: string,
): { projectType: ProjectAiProjectType; kind: ProjectAiFileKind } | null {
  const base = basename(name)
  for (const projectType of PROJECT_AI_PROJECT_TYPES) {
    for (const kind of ['doc', 'sheet', 'deck'] as const) {
      if (base === projectAiFileName(kind, projectType)) return { projectType, kind }
    }
  }
  return null
}

export function projectAiFileKindForName(name: string): ProjectAiFileKind | null {
  return projectAiFileInfoForName(name)?.kind ?? null
}

export function isProjectAiFileName(name: string): boolean {
  return projectAiFileKindForName(basename(name)) !== null
}
