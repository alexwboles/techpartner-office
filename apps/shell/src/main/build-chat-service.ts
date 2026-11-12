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
 * With no model configured, the deterministic fallback from the shared
 * core takes over: match one template, ask its placeholder questions,
 * fill from the answers. The panel says plainly when that mode is on.
 */

import { existsSync, readFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { basename, join } from 'node:path'
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
import { replaceInTextNodes } from './project-ai-service'
import { sanitizeFileBase, uniquePathIn } from './mcp/tools/document-tools'
import { buildTemplatePptx, buildTemplateXlsx } from './template-files'
import {
  applyBuildFill,
  buildChatSystemPrompt,
  describePlan,
  extractTemplatePlaceholders,
  fallbackCreateItem,
  fallbackQuestionsFor,
  isBuildConfirmation,
  matchTemplateForMessage,
  parseBuildReply,
  parsePlaceholderAnswers,
  planHasEdits,
  type BuildChatTurn,
  type BuildCreatedFile,
  type BuildPlanItem,
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

const CLARIFY_QUESTION =
  'Could you tell me a bit more about what you need — what should it be called, and who or what is it for?'

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
  /** questions the assistant asked on the last model turn (state only —
   *  the clarify gate is that a questions turn never executes) */
  private pendingQuestions: string[] = []
  private pendingEditPlan: BuildPlanItem[] | null = null
  private pendingFallback: PendingFallback | null = null

  constructor(deps: BuildChatServiceDeps) {
    this.deps = deps
  }

  modelAvailable(): boolean {
    return this.deps.planWithAi ? true : buildChatModelAvailable(this.deps.userDataPath)
  }

  reset(): void {
    this.history = []
    this.pendingQuestions = []
    this.pendingEditPlan = null
    this.pendingFallback = null
  }

  /** One user message in, one rendered turn out. */
  async send(rawText: string): Promise<BuildChatTurn> {
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
    return { ...done, planLines, modelAvailable: true }
  }

  private async respondWithoutModel(
    text: string,
    reason: 'not-configured' | 'unreachable',
  ): Promise<BuildChatTurn> {
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
    for (const item of items) {
      try {
        files.push(await this.executeItem(item))
      } catch (error) {
        const label =
          item.kind === 'edit' ? `Edit ${basename(item.path)}` : `Create “${item.title}”`
        errors.push(`${label}: ${error instanceof Error ? error.message : String(error)}`)
      }
    }
    const message = files.length > 0 ? intro : 'Nothing was created — see the problems below.'
    return turn({ kind: 'built', message, files, errors, modelAvailable: true })
  }

  private async executeItem(item: BuildPlanItem): Promise<BuildCreatedFile> {
    if (item.kind === 'edit') return this.executeEdit(item)
    const template = item.templateId
      ? TEMPLATES.find((candidate) => candidate.id === item.templateId)
      : undefined
    if (item.templateId && !template) throw new Error('That template no longer exists.')
    const filled = template ? applyBuildFill(template, item.fill ?? []) : undefined

    if (item.app === 'docs') {
      if (filled && filled.content.app === 'docs') {
        this.deps.openDocsTemplate({ title: item.title, html: filled.content.html })
        return { app: 'docs', title: item.title }
      }
      const markdown = item.markdown ?? ''
      const body = markdown.trimStart().startsWith('#')
        ? markdown
        : `# ${item.title}\n\n${markdown}`
      const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.docx`)
      await this.writeDocx(body, out)
      this.openQuietly(out)
      return { app: 'docs', title: item.title, path: out }
    }
    if (item.app === 'sheets') {
      const sheets = filled && filled.content.app === 'sheets' ? filled.content.sheets : item.sheets
      if (!sheets) throw new Error('The plan did not include any sheet content.')
      const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.xlsx`)
      await atomicWriteFile(out, await buildTemplateXlsx(sheets))
      this.openQuietly(out)
      return { app: 'sheets', title: item.title, path: out }
    }
    const slides = filled && filled.content.app === 'slides' ? filled.content.slides : item.slides
    if (!slides) throw new Error('The plan did not include any slides.')
    const out = uniquePathIn(this.deps.saveDir(), `${sanitizeFileBase(item.title)}.pptx`)
    await atomicWriteFile(out, await buildTemplatePptx(slides))
    this.openQuietly(out)
    return { app: 'slides', title: item.title, path: out }
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

  private pushHistory(role: 'user' | 'assistant', text: string): void {
    this.history.push({ role, text })
    if (this.history.length > HISTORY_LIMIT) {
      this.history = this.history.slice(this.history.length - HISTORY_LIMIT)
    }
  }

  private async callModel(system: string): Promise<string | null> {
    const target = resolveModelTarget(this.deps.userDataPath)
    if (!target) return null
    const transcript = this.history
      .map((entry) => `${entry.role === 'user' ? 'User' : 'Assistant'}: ${entry.text}`)
      .join('\n')
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), MODEL_TIMEOUT_MS)
    try {
      const result = await chatForProvider(
        target.provider,
        target.config,
        system,
        `Conversation so far:\n${transcript}\n\nReply with the JSON for the latest user message.`,
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
