/**
 * Build chat — shell main-process service.
 *
 * One conversation: the user describes what they need in plain words,
 * the model (the user's own configured provider — local Ollama by
 * default, or their key; never Genspark) replies with the JSON contract
 * from `../shared/build-chat`, and this service executes validated
 * plans with the app's real builders: gallery templates (Docs through
 * the AI-document pipeline, Sheets/Slides through template-files),
 * custom Docs through the bundled CLI, custom Sheets/Slides through
 * the same template builders.
 *
 * The gates, in order, every turn:
 * 1. A pending edit plan executes only when the user's next message is
 *    a plain confirmation; anything else drops the plan unexecuted.
 * 2. A questions reply never builds — the clarify gate is structural:
 *    execution happens only on a validated plan.
 * 3. An invalid model reply becomes a clarifying question, not a crash
 *    and not a guessed build.
 * 4. Creates land on unique-suffixed paths in the default save folder:
 *    a build can never overwrite an existing file.
 * 5. Edits (exact-text replacement in a document or deck) run only
 *    from a confirmed plan, and only when the text occurs exactly once.
 *
 * The rendered transcript is owned here and persisted to
 * build-chat-history.json in userData (last 200 entries), so the panel
 * hydrates on mount and the conversation survives view switches and
 * restarts. A pending edit plan is deliberately NOT persisted: on
 * load, a restored confirm card is marked expired, and no later "yes"
 * can release it — the gate state never outlives the process.
 *
 * With no model configured, the deterministic fallback from the shared
 * core takes over: match one template, ask its placeholder questions,
 * fill from the answers. The panel says plainly when that mode is on.
 */

import { existsSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, join, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import {
  activeProvider,
  chatForProvider,
  defaultAiSettings,
  resolveAiSettings,
} from '@genoffice/ai-provider'
import { atomicWriteFile } from './atomic-write'
import type { CliRunner } from './mcp/cli-runner'
import { createFileViaCli } from './mcp/headless-cli'
import { extractText } from './file-index/extract'
import { replaceInTextNodes } from './project-ai-service'
import { applySheetCellUpdates } from './sheet-edit'
import { sanitizeFileBase, uniquePathIn } from './mcp/tools/document-tools'
import { buildTemplatePptx, buildTemplateXlsx } from './template-files'
import {
  applyBuildFill,
  BUILD_CHAT_TRANSCRIPT_LIMIT,
  BUILD_DERIVE_TEXT_LIMIT,
  buildChatSystemPrompt,
  deriveDocumentSystemPrompt,
  describePlan,
  extractTemplatePlaceholders,
  fallbackCreateItem,
  fallbackQuestionsFor,
  isBuildConfirmation,
  matchTemplateForMessage,
  parseBuildReply,
  parsePlaceholderAnswers,
  planHasEdits,
  type BuildChatTranscriptEntry,
  type BuildChatTurn,
  type BuildCreatedFile,
  type BuildDeriveItem,
  type BuildPlanItem,
  type BuildSheetEditItem,
} from '../shared/build-chat'
import { TEMPLATES, type TemplateDef } from '../shared/template-defs'

export interface BuildChatHistoryEntry {
  role: 'user' | 'assistant'
  text: string
}

export interface BuildChatServiceDeps {
  userDataPath: string
  cli: CliRunner
  /** where new files land (the shell's default save folder) */
  saveDir: () => string
  /** open a created or edited file as a tab */
  openFile: (path: string) => void
  /** open a filled Docs template as a tab (the gallery's AI-document path);
   *  the file lands on disk under the title when the editor saves it */
  openDocsTemplate: (input: { title: string; html: string }) => void
  /** model override (tests inject a fake; production uses Ollama/BYOK) */
  planWithAi?: (input: {
    system: string
    history: BuildChatHistoryEntry[]
  }) => Promise<string | null>
  /** docx writer override (tests); production creates through the CLI */
  writeDocxFile?: (markdown: string, outPath: string) => Promise<void>
}

const HISTORY_LIMIT = 20
const MODEL_TIMEOUT_MS = 60_000

/** the persisted transcript file inside the app's userData folder */
const TRANSCRIPT_FILE = 'build-chat-history.json'

const EXPIRED_EDIT_NOTE = 'That edit plan expired when the app closed, so it was not run.'

const CLARIFY_QUESTION =
  'Could you tell me a bit more about what you need — what should it be called, and who or what is it for?'

/** A message that names an existing readable file — in the no-model
 *  fallback that can only be a derive request, which the fallback
 *  honestly cannot do (it builds from templates, it cannot read). */
const EXISTING_FILE_MENTION = /[\w][\w .-]*\.(docx|xlsx|pptx|md|html|txt)\b/i

function isRecordValue(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function storedStringList(value: unknown): string[] {
  if (!Array.isArray(value)) return []
  return value.filter((entry): entry is string => typeof entry === 'string')
}

/** Rebuild a turn from the persisted file, field by field: a hand-edited
 *  or half-written history file degrades to fewer entries, never to a
 *  crash or a malformed card. Returns undefined for a non-turn. */
function sanitizeStoredTurn(raw: unknown): BuildChatTurn | undefined {
  if (!isRecordValue(raw)) return undefined
  const kinds: Array<BuildChatTurn['kind']> = [
    'questions',
    'built',
    'confirm',
    'message',
    'no-model',
    'error',
  ]
  if (!kinds.includes(raw.kind as BuildChatTurn['kind'])) return undefined
  const files: BuildCreatedFile[] = []
  if (Array.isArray(raw.files)) {
    for (const entry of raw.files) {
      if (!isRecordValue(entry)) continue
      if (entry.app !== 'docs' && entry.app !== 'sheets' && entry.app !== 'slides') continue
      if (typeof entry.title !== 'string') continue
      files.push({
        app: entry.app,
        title: entry.title,
        ...(typeof entry.path === 'string' ? { path: entry.path } : {}),
      })
    }
  }
  return {
    kind: raw.kind as BuildChatTurn['kind'],
    message: typeof raw.message === 'string' ? raw.message : '',
    questions: storedStringList(raw.questions),
    planLines: storedStringList(raw.planLines),
    files,
    errors: storedStringList(raw.errors),
    modelAvailable: raw.modelAvailable === true,
  }
}

interface PendingFallback {
  template: TemplateDef
  placeholders: string[]
}

/** The model's JSON, fenced or bare; null when no JSON object is present. */
export function extractBuildJson(text: string): unknown | null {
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

/** Whether a usable model is configured (same resolution as Project AI:
 *  the user's own provider, Genspark never counts). */
export function buildChatModelAvailable(userDataPath: string): boolean {
  return resolveModelTarget(userDataPath) !== null
}

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
  // Never route a build request to Genspark: activeProvider already
  // migrates it to Custom, and this is a second guard.
  if (provider === 'genspark') return null
  const config = settings.providers?.[provider]
  if (!config) return null
  if (provider === 'custom' && !config.model?.trim()) return null
  return { provider, config }
}

function turn(partial: Partial<BuildChatTurn> & { kind: BuildChatTurn['kind'] }): BuildChatTurn {
  return {
    message: '',
    questions: [],
    planLines: [],
    files: [],
    errors: [],
    modelAvailable: true,
    ...partial,
  }
}

export class BuildChatService {
  private readonly deps: BuildChatServiceDeps
  private history: BuildChatHistoryEntry[] = []
  /** the rendered transcript: what the panel shows, persisted to disk */
  private transcriptEntries: BuildChatTranscriptEntry[] = []
  /** questions the assistant asked on the last model turn (state only —
   *  the clarify gate is that a questions turn never executes) */
  private pendingQuestions: string[] = []
  private pendingEditPlan: BuildPlanItem[] | null = null
  private pendingFallback: PendingFallback | null = null

  constructor(deps: BuildChatServiceDeps) {
    this.deps = deps
    this.loadTranscript()
  }

  modelAvailable(): boolean {
    return this.deps.planWithAi ? true : buildChatModelAvailable(this.deps.userDataPath)
  }

  /** The rendered transcript, oldest first — the panel hydrates from it. */
  transcript(): BuildChatTranscriptEntry[] {
    return JSON.parse(JSON.stringify(this.transcriptEntries)) as BuildChatTranscriptEntry[]
  }

  reset(): void {
    this.history = []
    this.pendingQuestions = []
    this.pendingEditPlan = null
    this.pendingFallback = null
    this.transcriptEntries = []
    try {
      rmSync(join(this.deps.userDataPath, TRANSCRIPT_FILE), { force: true })
    } catch {
      // a stuck file must not break the reset; the next save overwrites it
    }
  }

  /** One user message in, one rendered turn out. Every exchange is
   *  recorded in the transcript (and persisted) whatever its outcome. */
  async send(rawText: string): Promise<BuildChatTurn> {
    const result = await this.takeTurn(rawText)
    const text = rawText.trim()
    if (text) {
      this.transcriptEntries.push(
        { role: 'user', text },
        { role: 'assistant', text: result.message, turn: result },
      )
      if (this.transcriptEntries.length > BUILD_CHAT_TRANSCRIPT_LIMIT) {
        this.transcriptEntries = this.transcriptEntries.slice(-BUILD_CHAT_TRANSCRIPT_LIMIT)
      }
      this.saveTranscript()
    }
    return result
  }

  private async takeTurn(rawText: string): Promise<BuildChatTurn> {
    const text = rawText.trim()
    const modelAvailable = this.modelAvailable()
    if (!text) {
      return turn({
        kind: 'error',
        message: 'Type what you need first — for example, “an invoice for my consulting work”.',
        modelAvailable,
      })
    }
    this.pushHistory('user', text)

    // Gate 1: a pending edit plan lives or dies by this message.
    if (this.pendingEditPlan) {
      const plan = this.pendingEditPlan
      this.pendingEditPlan = null
      if (isBuildConfirmation(text)) {
        const done = await this.executePlan(plan, 'Done — here is what changed:')
        this.pushHistory('assistant', done.message)
        return { ...done, modelAvailable }
      }
      const next = await this.respond(text, modelAvailable)
      return {
        ...next,
        message: next.message
          ? `Dropped that edit plan. ${next.message}`
          : 'Dropped that edit plan.',
      }
    }

    // The deterministic fallback is mid-questions: this message is the
    // answer set, not a new request.
    if (this.pendingFallback) {
      const pending = this.pendingFallback
      const answers = parsePlaceholderAnswers(pending.placeholders, text)
      if (!answers) {
        return turn({
          kind: 'questions',
          message:
            'I still need those details before I can build it — one answer per line, in the order I asked, or label each one like “[Name]: value”.',
          questions: fallbackQuestionsFor(pending.placeholders),
          modelAvailable,
        })
      }
      this.pendingFallback = null
      const item = fallbackCreateItem(pending.template, answers)
      const done = await this.executePlan(
        [item],
        `Built from the “${pending.template.title}” template:`,
      )
      this.pushHistory('assistant', done.message)
      return { ...done, modelAvailable }
    }

    const next = await this.respond(text, modelAvailable)
    if (next.kind === 'built' || next.kind === 'questions' || next.kind === 'confirm') {
      this.pushHistory('assistant', next.message || next.planLines.join('; '))
    }
    return next
  }

  private async respond(text: string, modelAvailable: boolean): Promise<BuildChatTurn> {
    if (modelAvailable) return this.respondWithModel(text)
    return this.respondWithoutModel(text, 'not-configured')
  }

  private async respondWithModel(text: string): Promise<BuildChatTurn> {
    const system = buildChatSystemPrompt(TEMPLATES)
    const replyText = this.deps.planWithAi
      ? await this.deps.planWithAi({ system, history: [...this.history] })
      : await this.callModel(system)
    if (replyText === null) {
      // A configured model that does not answer (local AI not running,
      // a bad key) degrades to the template path, never to a dead end —
      // the same degrade Project AI uses for its drafts.
      return this.respondWithoutModel(text, 'unreachable')
    }
    const parsed = parseBuildReply(extractBuildJson(replyText), TEMPLATES)
    if (!parsed) {
      // Gate 3: unparseable model output becomes a clarifying question
      this.pendingQuestions = [CLARIFY_QUESTION]
      return turn({
        kind: 'questions',
        message: 'I want to make sure I build the right thing.',
        questions: [CLARIFY_QUESTION],
        modelAvailable: true,
      })
    }
    if (parsed.kind === 'questions') {
      // Gate 2: a questions turn never executes anything
      this.pendingQuestions = parsed.questions
      return turn({
        kind: 'questions',
        message: parsed.message ?? 'A few details first, so the result is right:',
        questions: parsed.questions,
        modelAvailable: true,
      })
    }
    this.pendingQuestions = []
    const planLines = describePlan(parsed.items, TEMPLATES)
    if (planHasEdits(parsed.items)) {
      // Gate 5: edits wait for an explicit yes in the next message
      this.pendingEditPlan = parsed.items
      return turn({
        kind: 'confirm',
        message:
          parsed.message ??
          'Here is the plan. It changes an existing file — reply “yes” and I will run it, or say what to change instead.',
        planLines,
        modelAvailable: true,
      })
    }
    const done = await this.executePlan(parsed.items, parsed.message ?? 'Done — created:')
    // execution can add its own plan lines (a derive that only used the
    // first part of a long source says so in the plan card)
    return { ...done, planLines: [...planLines, ...done.planLines], modelAvailable: true }
  }

  private async respondWithoutModel(
    text: string,
    reason: 'not-configured' | 'unreachable',
  ): Promise<BuildChatTurn> {
    if (EXISTING_FILE_MENTION.test(text)) {
      // Working from an existing file needs a model that can read it;
      // say so plainly instead of template-matching around the request.
      return turn({
        kind: 'no-model',
        message:
          reason === 'not-configured'
            ? 'That asks me to work from an existing file, which needs AI — and no AI is set up yet. Set up AI (free local Ollama, or your own key) in Settings, then ask again. Without AI I can only build straight from a template.'
            : 'That asks me to work from an existing file, which needs your AI — and it did not answer. Check that your AI is running (for local AI, that Ollama is up), then try again.',
        modelAvailable: reason === 'not-configured' ? false : true,
      })
    }
    const template = matchTemplateForMessage(text, TEMPLATES)
    if (!template) {
      return turn({
        kind: 'no-model',
        message:
          reason === 'not-configured'
            ? 'No AI is set up yet, so I can only build straight from a template here. Name what you need — an invoice, a resume, a budget, a letter — and I will match a template and ask for its details. For full chat building, set up AI (free local Ollama, or your own key) in Settings.'
            : 'Your AI did not answer, and that request does not match a template I can build directly. Check that your AI is running (for local AI, that Ollama is up), then try again — or name a template, like an invoice or a resume.',
        modelAvailable: reason === 'not-configured' ? false : true,
      })
    }
    const placeholders = extractTemplatePlaceholders(template)
    if (placeholders.length > 0) {
      this.pendingFallback = { template, placeholders }
      return turn({
        kind: 'questions',
        message:
          reason === 'not-configured'
            ? `No AI is set up, so I matched the “${template.title}” template. I need its details before I build it:`
            : `Your AI did not answer, so I matched the “${template.title}” template instead. I need its details before I build it:`,
        questions: fallbackQuestionsFor(placeholders),
        modelAvailable: reason === 'not-configured' ? false : true,
      })
    }
    const done = await this.executePlan(
      [fallbackCreateItem(template, null)],
      `Built from the “${template.title}” template:`,
    )
    return { ...done, modelAvailable: false }
  }

  /** Execute a validated plan, item by item; a failed item is reported,
   *  never silently skipped, and never stops the rest of the plan. */
  private async executePlan(items: BuildPlanItem[], intro: string): Promise<BuildChatTurn> {
    const files: BuildCreatedFile[] = []
    const errors: string[] = []
    const notes: string[] = []
    for (const item of items) {
      try {
        const result = await this.executeItem(item)
        files.push(result.file)
        if (result.note) notes.push(result.note)
      } catch (error) {
        const label =
          item.kind === 'create' || item.kind === 'derive'
            ? `Create “${item.title}”`
            : `Edit ${basename(item.path)}`
        errors.push(`${label}: ${error instanceof Error ? error.message : String(error)}`)
      }
    }
    const message = files.length > 0 ? intro : 'Nothing was created — see the problems below.'
    return turn({ kind: 'built', message, files, errors, planLines: notes, modelAvailable: true })
  }

  private async executeItem(
    item: BuildPlanItem,
  ): Promise<{ file: BuildCreatedFile; note?: string }> {
    if (item.kind === 'edit') return { file: await this.executeEdit(item) }
    if (item.kind === 'sheet-edit') return { file: await this.executeSheetEdit(item) }
    if (item.kind === 'derive') return this.executeDerive(item)
    const template = item.templateId
      ? TEMPLATES.find((candidate) => candidate.id === item.templateId)
      : undefined
    if (item.templateId && !template) throw new Error('That template no longer exists.')
    const filled = template ? applyBuildFill(template, item.fill ?? []) : undefined

    if (item.app === 'docs') {
      if (filled && filled.content.app === 'docs') {
        this.deps.openDocsTemplate({ title: item.title, html: filled.content.html })
        return { file: { app: 'docs', title: item.title } }
      }
      const markdown = item.markdown ?? ''
      const body = markdown.trimStart().startsWith('#')
        ? markdown
        : `# ${item.title}\n\n${markdown}`
      const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.docx`)
      await this.writeDocx(body, out)
      this.openQuietly(out)
      return { file: { app: 'docs', title: item.title, path: out } }
    }
    if (item.app === 'sheets') {
      const sheets = filled && filled.content.app === 'sheets' ? filled.content.sheets : item.sheets
      if (!sheets) throw new Error('The plan did not include any sheet content.')
      const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.xlsx`)
      await atomicWriteFile(out, await buildTemplateXlsx(sheets))
      this.openQuietly(out)
      return { file: { app: 'sheets', title: item.title, path: out } }
    }
    const slides = filled && filled.content.app === 'slides' ? filled.content.slides : item.slides
    if (!slides) throw new Error('The plan did not include any slides.')
    const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.pptx`)
    await atomicWriteFile(out, await buildTemplatePptx(slides))
    this.openQuietly(out)
    return { file: { app: 'slides', title: item.title, path: out } }
  }

  /** Exact-text edit of an existing document or deck (confirmed plans
   *  only): the text must occur exactly once across the file's text
   *  parts, the same rule Project AI applies to its queued changes. */
  private async executeEdit(
    item: Extract<BuildPlanItem, { kind: 'edit' }>,
  ): Promise<BuildCreatedFile> {
    if (!existsSync(item.path)) throw new Error('That file does not exist.')
    const bytes = await readFile(item.path)
    const zip = await JSZip.loadAsync(bytes)
    const targets =
      item.app === 'docs'
        ? ['word/document.xml']
        : Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    let changed = false
    for (const name of targets) {
      const entry = zip.files[name]
      if (!entry) continue
      const xml = await entry.async('string')
      const next = replaceInTextNodes(xml, item.find, item.replace, false)
      if (next !== null && next !== xml) {
        zip.file(name, next)
        changed = true
        break
      }
    }
    if (!changed) {
      throw new Error('That text was not found exactly once, so nothing was changed.')
    }
    const out = await zip.generateAsync({ type: 'nodebuffer' })
    await atomicWriteFile(item.path, out)
    this.openQuietly(item.path)
    return { app: item.app, title: basename(item.path), path: item.path }
  }

  /** Resolve a save-folder-relative plan path, refusing anything that
   *  lands outside the folder (the validator already rejects absolute
   *  paths and ".." segments; this is the second, physical check). */
  private resolveSaveFolderPath(relative: string): string {
    const root = resolve(this.deps.saveDir())
    const target = resolve(root, relative)
    if (target !== root && !target.startsWith(root + sep)) {
      throw new Error('That file is outside the save folder.')
    }
    return target
  }

  /** Literal cell-value edit of an existing spreadsheet (confirmed
   *  plans only): values land exactly as listed, the rest of the
   *  workbook is preserved by the sheet-edit helper. */
  private async executeSheetEdit(item: BuildSheetEditItem): Promise<BuildCreatedFile> {
    const path = this.resolveSaveFolderPath(item.path)
    if (!existsSync(path)) throw new Error('That file does not exist.')
    const bytes = await readFile(path)
    const out = await applySheetCellUpdates(bytes, item.updates, item.sheet)
    await atomicWriteFile(path, out)
    this.openQuietly(path)
    return { app: 'sheets', title: basename(path), path }
  }

  /** Derive: read an existing file's text (the same extractors Home
   *  search uses), hand it to the model with the user's instruction,
   *  and create a NEW document from what comes back. The source is only
   *  ever read. Long sources are cut at BUILD_DERIVE_TEXT_LIMIT and the
   *  plan card says so — the model never silently sees a partial file
   *  presented as the whole. */
  private async executeDerive(
    item: BuildDeriveItem,
  ): Promise<{ file: BuildCreatedFile; note?: string }> {
    const sourcePath = this.resolveSaveFolderPath(item.source)
    if (!existsSync(sourcePath)) throw new Error('That source file does not exist.')
    const extracted = await extractText(sourcePath)
    if (extracted.kind !== 'text' || !extracted.text.trim()) {
      throw new Error('I could not read any text from that file.')
    }
    const truncated = extracted.text.length > BUILD_DERIVE_TEXT_LIMIT
    const content = truncated ? extracted.text.slice(0, BUILD_DERIVE_TEXT_LIMIT) : extracted.text
    const markdown = await this.deriveMarkdown(item, content, truncated)
    if (!markdown) throw new Error('The AI did not produce a document from that file.')
    const body = markdown.trimStart().startsWith('#') ? markdown : `# ${item.title}\n\n${markdown}`
    const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.docx`)
    await this.writeDocx(body, out)
    this.openQuietly(out)
    return {
      file: { app: 'docs', title: item.title, path: out },
      ...(truncated
        ? {
            note:
              `“${basename(item.source)}” is long — ` +
              'the new document is based on the first part of the file.',
          }
        : {}),
    }
  }

  /** The derive content call: the writing prompt, not the plan
   *  contract. A fenced reply is unfenced; an empty one is a failure. */
  private async deriveMarkdown(
    item: BuildDeriveItem,
    content: string,
    truncated: boolean,
  ): Promise<string | null> {
    const system = deriveDocumentSystemPrompt()
    const userText =
      `Instruction: ${item.instruction}\n\n` +
      `Source file: ${basename(item.source)}${truncated ? ' (only the first part is shown)' : ''}\n\n` +
      content
    const reply = this.deps.planWithAi
      ? await this.deps.planWithAi({ system, history: [{ role: 'user', text: userText }] })
      : await this.callModelRaw(system, userText)
    if (!reply) return null
    const unfenced = reply
      .replace(/^\s*```(?:markdown|md)?\s*/i, '')
      .replace(/\s*```\s*$/, '')
      .trim()
    return unfenced ? unfenced : null
  }

  private async writeDocx(markdown: string, out: string): Promise<void> {
    if (this.deps.writeDocxFile) {
      await this.deps.writeDocxFile(markdown, out)
      return
    }
    await createFileViaCli(this.deps.cli, {
      type: 'docx',
      input: { name: 'content.md', content: markdown },
      out,
      overwrite: false,
    })
  }

  private openQuietly(path: string): void {
    try {
      this.deps.openFile(path)
    } catch {
      // opening a tab must not fail a build; the file is on disk
    }
  }

  /** Restore the persisted transcript (constructor only). Entries are
   *  rebuilt field by field; a corrupt file means a fresh conversation.
   *  Restored confirm cards are marked expired — their plan lived only
   *  in the old process's memory and is gone for good. The model history
   *  is rebuilt from the same entries so the conversation continues
   *  coherently (a pending question set still reads as asked). */
  private loadTranscript(): void {
    let parsed: unknown
    try {
      parsed = JSON.parse(
        readFileSync(join(this.deps.userDataPath, TRANSCRIPT_FILE), 'utf8'),
      ) as unknown
    } catch {
      return
    }
    if (!isRecordValue(parsed) || !Array.isArray(parsed.entries)) return
    const entries: BuildChatTranscriptEntry[] = []
    for (const raw of parsed.entries) {
      if (!isRecordValue(raw)) continue
      if (raw.role !== 'user' && raw.role !== 'assistant') continue
      if (typeof raw.text !== 'string') continue
      const storedTurn = raw.role === 'assistant' ? sanitizeStoredTurn(raw.turn) : undefined
      entries.push({
        role: raw.role,
        text: raw.text,
        ...(storedTurn ? { turn: storedTurn } : {}),
      })
    }
    this.transcriptEntries = entries.slice(-BUILD_CHAT_TRANSCRIPT_LIMIT)
    for (const entry of this.transcriptEntries) {
      if (entry.turn?.kind === 'confirm') {
        const message = [entry.turn.message, EXPIRED_EDIT_NOTE].filter(Boolean).join(' ')
        entry.turn = { ...entry.turn, kind: 'message', message }
        entry.text = message
      }
      if (entry.role === 'user') this.pushHistory('user', entry.text)
      else if (entry.text) this.pushHistory('assistant', entry.text)
    }
  }

  /** Persist the transcript atomically (temp file + rename, the same
   *  shape as the app's other small JSON stores). Best-effort: a failed
   *  write loses the saved transcript, never the live conversation. */
  private saveTranscript(): void {
    const file = join(this.deps.userDataPath, TRANSCRIPT_FILE)
    const tmp = `${file}.tmp`
    try {
      writeFileSync(tmp, JSON.stringify({ version: 1, entries: this.transcriptEntries }), 'utf8')
      renameSync(tmp, file)
    } catch {
      // ignored — see above
    }
  }

  private pushHistory(role: 'user' | 'assistant', text: string): void {
    this.history.push({ role, text })
    if (this.history.length > HISTORY_LIMIT) {
      this.history = this.history.slice(this.history.length - HISTORY_LIMIT)
    }
  }

  private async callModel(system: string): Promise<string | null> {
    const transcript = this.history
      .map((entry) => `${entry.role === 'user' ? 'User' : 'Assistant'}: ${entry.text}`)
      .join('\n')
    return this.callModelRaw(
      system,
      `Conversation so far:\n${transcript}\n\nReply with the JSON for the latest user message.`,
    )
  }

  /** One raw model call against the configured provider (never
   *  Genspark), with the shared timeout; null when it cannot answer. */
  private async callModelRaw(system: string, prompt: string): Promise<string | null> {
    const target = resolveModelTarget(this.deps.userDataPath)
    if (!target) return null
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), MODEL_TIMEOUT_MS)
    try {
      const result = await chatForProvider(
        target.provider,
        target.config,
        system,
        prompt,
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
}
