/**
 * Build chat — pure, framework-free core.
 *
 * The Build chat in Home turns a plain-words request into real files:
 * the model (the user's own configured provider — local Ollama or their
 * key) replies with JSON only, in exactly one of two shapes:
 *
 *   { "questions": ["...", ...] }                  — facts are missing
 *   { "plan": [ <item>, ... ], "message": "..." }  — ready to build
 *
 * A plan item either creates a NEW file (from a gallery template with
 * the user's facts filled in, or from custom content) or — only for a
 * document or a deck — edits an existing file by exact text replacement.
 * This module validates every reply before anything touches disk: an
 * invalid reply parses to null and the service asks a clarifying
 * question instead of crashing or half-building.
 *
 * SAFETY CONTRACT:
 * - Creates never overwrite: the service writes to a unique-suffixed
 *   path, and this module has no notion of an overwrite flag at all.
 * - Edits are a separate item kind, docs/slides only, exact-text only,
 *   and the service executes them only after the user confirms in chat.
 * - The model never invents facts silently: the system prompt below
 *   tells it to ask when a name, amount, date or item that would end
 *   up in a file is missing, and `parseBuildReply` cannot turn a
 *   questions reply into a build — the two shapes are disjoint.
 *
 * With no model configured, a small deterministic fallback lives here
 * too: match the request to one template by word overlap, ask for that
 * template's [bracketed placeholders], and fill them from the answers.
 * It is deliberately narrow (template pick + fill only) and the panel
 * says plainly when it is in that mode.
 *
 * This module runs in the shell main process, the shell renderer, and
 * vitest. It must not import node or Electron APIs.
 */

import type { TemplateDef, TemplateSheet, TemplateSheetCell, TemplateSlide } from './template-defs'

export type BuildChatApp = 'docs' | 'sheets' | 'slides'

export interface BuildFillEntry {
  /** exact text to find (usually a [bracketed placeholder]) */
  find: string
  /** the user's fact that replaces it */
  replace: string
}

export interface BuildCreateItem {
  kind: 'create'
  app: BuildChatApp
  title: string
  /** build from this gallery template (fill applies on top of it) */
  templateId?: string
  fill?: BuildFillEntry[]
  /** custom Docs content (markdown), when no template is used */
  markdown?: string
  /** custom Sheets content, when no template is used */
  sheets?: TemplateSheet[]
  /** custom Slides content, when no template is used */
  slides?: TemplateSlide[]
}

export interface BuildEditItem {
  kind: 'edit'
  app: 'docs' | 'slides'
  /** absolute path of the existing file to edit */
  path: string
  /** exact text to find (must occur exactly once in the file's text) */
  find: string
  /** replacement text */
  replace: string
}

export type BuildPlanItem = BuildCreateItem | BuildEditItem

export type BuildReply =
  | { kind: 'questions'; message?: string; questions: string[] }
  | { kind: 'plan'; message?: string; items: BuildPlanItem[] }

/** One file a build produced; `path` is absent for a Docs template,
 *  which opens as a tab and lands on disk when the editor saves it. */
export interface BuildCreatedFile {
  app: BuildChatApp
  title: string
  path?: string
}

/** One service turn, as rendered by the panel. */
export interface BuildChatTurn {
  kind: 'questions' | 'built' | 'confirm' | 'message' | 'no-model' | 'error'
  message: string
  questions: string[]
  planLines: string[]
  files: BuildCreatedFile[]
  errors: string[]
  modelAvailable: boolean
}

export const BUILD_CHAT_MAX_QUESTIONS = 5
export const BUILD_CHAT_MAX_ITEMS = 6
export const BUILD_CHAT_MAX_PLACEHOLDERS = 8

const MAX_QUESTION_CHARS = 300
const MAX_MESSAGE_CHARS = 1000
const MAX_TITLE_CHARS = 120
const MAX_MARKDOWN_CHARS = 30_000
const MAX_FILL_ENTRIES = 20
const MAX_FIND_CHARS = 200
const MAX_REPLACE_CHARS = 2000
const MAX_SHEETS = 3
const MAX_ROWS_PER_SHEET = 200
const MAX_CELLS_PER_ROW = 16
const MAX_CELL_CHARS = 500
const MAX_SLIDES = 20
const MAX_BULLETS_PER_SLIDE = 12

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function cleanString(value: unknown, maxChars: number): string {
  return typeof value === 'string' ? value.trim().slice(0, maxChars) : ''
}

function optionalMessage(value: unknown): string | undefined {
  const message = cleanString(value, MAX_MESSAGE_CHARS)
  return message ? message : undefined
}

function isBuildApp(value: unknown): value is BuildChatApp {
  return value === 'docs' || value === 'sheets' || value === 'slides'
}

/** A path the edit gate accepts: absolute POSIX or Windows form. */
export function isAbsoluteBuildPath(value: string): boolean {
  if (value.startsWith('/')) return true
  return /^[A-Za-z]:[\\/]/.test(value)
}

function validateFill(value: unknown): BuildFillEntry[] | null {
  if (value === undefined) return []
  if (!Array.isArray(value)) return null
  const out: BuildFillEntry[] = []
  for (const entry of value.slice(0, MAX_FILL_ENTRIES)) {
    if (!isRecord(entry)) return null
    const find = typeof entry.find === 'string' ? entry.find : ''
    const replace = typeof entry.replace === 'string' ? entry.replace : ''
    if (!find || find.length > MAX_FIND_CHARS || replace.length > MAX_REPLACE_CHARS) return null
    out.push({ find, replace })
  }
  return out
}

function validateCell(value: unknown): TemplateSheetCell | null {
  if (typeof value === 'string') return value.length <= MAX_CELL_CHARS ? value : null
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (isRecord(value)) {
    const formula = cleanString(value.formula, MAX_CELL_CHARS)
    const cached = value.cached
    if (!formula) return null
    if (typeof cached !== 'number' && typeof cached !== 'string') return null
    if (typeof cached === 'number' && !Number.isFinite(cached)) return null
    if (typeof cached === 'string' && cached.length > MAX_CELL_CHARS) return null
    const cell: {
      formula: string
      cached: number | string
      format?: 'currency' | 'percent' | 'int'
      bold?: boolean
    } = {
      formula,
      cached,
    }
    if (value.format === 'currency' || value.format === 'percent' || value.format === 'int') {
      cell.format = value.format
    }
    if (value.bold === true) cell.bold = true
    return cell
  }
  return null
}

const SHEET_NAME_FORBIDDEN = '[]:*?/\\'

function validateSheets(value: unknown): TemplateSheet[] | null {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_SHEETS) return null
  const out: TemplateSheet[] = []
  for (const entry of value) {
    if (!isRecord(entry)) return null
    const name = cleanString(entry.name, 31)
    if (!name) return null
    if ([...name].some((ch) => SHEET_NAME_FORBIDDEN.includes(ch))) return null
    if (!Array.isArray(entry.rows) || entry.rows.length === 0) return null
    if (entry.rows.length > MAX_ROWS_PER_SHEET) return null
    const rows: TemplateSheetCell[][] = []
    for (const rawRow of entry.rows) {
      if (!Array.isArray(rawRow) || rawRow.length > MAX_CELLS_PER_ROW) return null
      const row: TemplateSheetCell[] = []
      for (const rawCell of rawRow) {
        const cell = validateCell(rawCell)
        if (cell === null) return null
        row.push(cell)
      }
      rows.push(row)
    }
    const sheet: TemplateSheet = { name, rows }
    if (entry.headerRow === true) sheet.headerRow = true
    out.push(sheet)
  }
  return out
}

function validateSlides(value: unknown): TemplateSlide[] | null {
  if (!Array.isArray(value) || value.length === 0 || value.length > MAX_SLIDES) return null
  const out: TemplateSlide[] = []
  for (const entry of value) {
    if (!isRecord(entry)) return null
    const title = cleanString(entry.title, 200)
    if (!title) return null
    const bulletsRaw = entry.bullets === undefined ? [] : entry.bullets
    if (!Array.isArray(bulletsRaw) || bulletsRaw.length > MAX_BULLETS_PER_SLIDE) return null
    const bullets: string[] = []
    for (const bullet of bulletsRaw) {
      const text = cleanString(bullet, 400)
      if (!text) return null
      bullets.push(text)
    }
    const slide: TemplateSlide = { title, bullets }
    const subtitle = cleanString(entry.subtitle, 300)
    if (subtitle) slide.subtitle = subtitle
    out.push(slide)
  }
  return out
}

function validateCreateItem(
  entry: Record<string, unknown>,
  templates: readonly TemplateDef[],
): BuildCreateItem | null {
  if (!isBuildApp(entry.app)) return null
  const app = entry.app
  const title = cleanString(entry.title, MAX_TITLE_CHARS)
  if (!title) return null
  const fill = validateFill(entry.fill)
  if (fill === null) return null
  const item: BuildCreateItem = { kind: 'create', app, title }
  if (fill.length > 0) item.fill = fill
  if (typeof entry.templateId === 'string' && entry.templateId) {
    const template = templates.find((candidate) => candidate.id === entry.templateId)
    // an unknown template, or one for another app, invalidates the reply:
    // building something else than the plan said would be a silent swap
    if (!template || template.content.app !== app) return null
    item.templateId = template.id
    return item
  }
  // custom content: exactly the payload for the item's app
  if (app === 'docs') {
    const markdown = typeof entry.markdown === 'string' ? entry.markdown : ''
    if (!markdown.trim() || markdown.length > MAX_MARKDOWN_CHARS) return null
    item.markdown = markdown
    return item
  }
  if (app === 'sheets') {
    const sheets = validateSheets(entry.sheets)
    if (!sheets) return null
    item.sheets = sheets
    return item
  }
  const slides = validateSlides(entry.slides)
  if (!slides) return null
  item.slides = slides
  return item
}

function validateEditItem(entry: Record<string, unknown>): BuildEditItem | null {
  if (entry.app !== 'docs' && entry.app !== 'slides') return null
  const path = typeof entry.path === 'string' ? entry.path.trim() : ''
  const find = typeof entry.find === 'string' ? entry.find : ''
  const replace = typeof entry.replace === 'string' ? entry.replace : ''
  if (!path || !isAbsoluteBuildPath(path)) return null
  if (!find || find.length > MAX_FIND_CHARS) return null
  if (replace.length > MAX_REPLACE_CHARS || find === replace) return null
  return { kind: 'edit', app: entry.app, path, find, replace }
}

/**
 * Validate a model reply into a BuildReply. Returns null for anything
 * that is not exactly one of the two contract shapes — the caller then
 * asks a clarifying question rather than guessing at the model's intent.
 */
export function parseBuildReply(
  value: unknown,
  templates: readonly TemplateDef[],
): BuildReply | null {
  if (!isRecord(value)) return null
  const message = optionalMessage(value.message)
  if (Array.isArray(value.questions)) {
    const questions: string[] = []
    for (const raw of value.questions) {
      const question = cleanString(raw, MAX_QUESTION_CHARS)
      if (question) questions.push(question)
    }
    if (questions.length === 0) return null
    return {
      kind: 'questions',
      ...(message ? { message } : {}),
      questions: questions.slice(0, BUILD_CHAT_MAX_QUESTIONS),
    }
  }
  if (Array.isArray(value.plan)) {
    if (value.plan.length === 0 || value.plan.length > BUILD_CHAT_MAX_ITEMS) return null
    const items: BuildPlanItem[] = []
    for (const raw of value.plan) {
      if (!isRecord(raw)) return null
      const item = raw.kind === 'edit' ? validateEditItem(raw) : validateCreateItem(raw, templates)
      if (!item) return null
      items.push(item)
    }
    return {
      kind: 'plan',
      ...(message ? { message } : {}),
      items,
    }
  }
  return null
}

/** True when a plan touches an existing file — those plans always wait
 *  for the user's explicit confirmation before anything executes. */
export function planHasEdits(items: readonly BuildPlanItem[]): boolean {
  return items.some((item) => item.kind === 'edit')
}

/** A short yes in chat: the only thing that releases a pending edit plan. */
export function isBuildConfirmation(text: string): boolean {
  const normalized = text
    .trim()
    .toLowerCase()
    .replace(/[.!]+$/, '')
  if (!normalized) return false
  const words = normalized.split(/\s+/)
  const first = words[0]?.replace(/[,;:]+$/, '') ?? ''
  if (['yes', 'y', 'yeah', 'yep', 'confirm', 'ok', 'okay'].includes(first)) {
    return words.length <= 4
  }
  return normalized === 'go ahead' || normalized === 'do it'
}

const APP_NOUN: Record<BuildChatApp, string> = {
  docs: 'document',
  sheets: 'spreadsheet',
  slides: 'presentation',
}

function truncateMiddle(text: string, max: number): string {
  return text.length > max ? `${text.slice(0, max - 1)}…` : text
}

function baseName(path: string): string {
  const parts = path.split(/[/\\]/)
  return parts[parts.length - 1] || path
}

/** One plain-language line per plan item — the plan the user sees in the
 *  chat before any file is created. */
export function describePlanItem(item: BuildPlanItem, templates: readonly TemplateDef[]): string {
  if (item.kind === 'edit') {
    return (
      `Edit ${baseName(item.path)}: replace “${truncateMiddle(item.find, 60)}” ` +
      `with “${truncateMiddle(item.replace, 60)}”`
    )
  }
  const noun = APP_NOUN[item.app]
  if (item.templateId) {
    const template = templates.find((candidate) => candidate.id === item.templateId)
    const templateTitle = template ? template.title : item.templateId
    return `Create a ${noun} “${item.title}” from the “${templateTitle}” template`
  }
  return `Create a ${noun} “${item.title}”`
}

export function describePlan(
  items: readonly BuildPlanItem[],
  templates: readonly TemplateDef[],
): string[] {
  return items.map((item) => describePlanItem(item, templates))
}

// ── Templates: placeholders + fill ────────────────────────────────────

const PLACEHOLDER_PATTERN = /\[[^\[\]\n]{1,80}\]/g

function placeholdersIn(text: string, into: string[]): void {
  for (const match of text.matchAll(PLACEHOLDER_PATTERN)) {
    const token = match[0]
    // a placeholder names something to fill in; a token with no letter
    // (a checkbox like [ ], a footnote marker) is not one
    if (!/[A-Za-z]/.test(token)) continue
    if (!into.includes(token)) into.push(token)
  }
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, ' ')
}

/**
 * The [bracketed placeholders] a template asks the user to replace, in
 * document order, de-duplicated. The deterministic fallback asks one
 * question per placeholder; the model path uses them as fill targets.
 */
export function extractTemplatePlaceholders(template: TemplateDef): string[] {
  const out: string[] = []
  const content = template.content
  if (content.app === 'docs') {
    placeholdersIn(stripTags(content.html), out)
  } else if (content.app === 'sheets') {
    for (const sheet of content.sheets) {
      placeholdersIn(sheet.name, out)
      for (const row of sheet.rows) {
        for (const cell of row) {
          if (typeof cell === 'string') placeholdersIn(cell, out)
        }
      }
    }
  } else {
    for (const slide of content.slides) {
      placeholdersIn(slide.title, out)
      if (slide.subtitle) placeholdersIn(slide.subtitle, out)
      for (const bullet of slide.bullets) placeholdersIn(bullet, out)
    }
  }
  return out.slice(0, BUILD_CHAT_MAX_PLACEHOLDERS)
}

function applyToText(text: string, fill: readonly BuildFillEntry[]): string {
  let out = text
  for (const entry of fill) {
    if (entry.find) out = out.split(entry.find).join(entry.replace)
  }
  return out
}

function applyToCell(cell: TemplateSheetCell, fill: readonly BuildFillEntry[]): TemplateSheetCell {
  if (typeof cell === 'string') return applyToText(cell, fill)
  if (typeof cell === 'object' && typeof cell.cached === 'string') {
    return { ...cell, cached: applyToText(cell.cached, fill) }
  }
  return cell
}

/**
 * A copy of a template with the user's facts filled in: every `find`
 * text replaced by its `replace` value across the document HTML, sheet
 * names and string cells, and slide titles/subtitles/bullets. The input
 * template is never mutated, and formula strings are left alone (their
 * cached string results do get the fill, the way a reader sees them).
 */
export function applyBuildFill(
  template: TemplateDef,
  fill: readonly BuildFillEntry[],
): TemplateDef {
  if (fill.length === 0) return template
  const content = template.content
  if (content.app === 'docs') {
    return { ...template, content: { app: 'docs', html: applyToText(content.html, fill) } }
  }
  if (content.app === 'sheets') {
    return {
      ...template,
      content: {
        app: 'sheets',
        sheets: content.sheets.map((sheet) => ({
          ...sheet,
          name: applyToText(sheet.name, fill),
          rows: sheet.rows.map((row) => row.map((cell) => applyToCell(cell, fill))),
        })),
      },
    }
  }
  return {
    ...template,
    content: {
      app: 'slides',
      slides: content.slides.map((slide) => ({
        ...slide,
        title: applyToText(slide.title, fill),
        ...(slide.subtitle ? { subtitle: applyToText(slide.subtitle, fill) } : {}),
        bullets: slide.bullets.map((bullet) => applyToText(bullet, fill)),
      })),
    },
  }
}

// ── Deterministic fallback (no model configured) ─────────────────────

function wordsOf(text: string): string[] {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter((word) => word.length >= 3)
}

/**
 * The one template a model-free request clearly points at, by word
 * overlap with template titles (weighted double), descriptions and
 * categories. Returns null when nothing scores or two templates tie —
 * guessing between templates would build the wrong file.
 */
export function matchTemplateForMessage(
  message: string,
  templates: readonly TemplateDef[],
): TemplateDef | null {
  const messageWords = new Set(wordsOf(message))
  if (messageWords.size === 0) return null
  let best: TemplateDef | null = null
  let bestScore = 0
  let secondScore = 0
  for (const template of templates) {
    const titleWords = new Set(wordsOf(template.title))
    const otherWords = new Set(wordsOf(`${template.description} ${template.category}`))
    let score = 0
    for (const word of messageWords) {
      if (titleWords.has(word)) score += 2
      else if (otherWords.has(word)) score += 1
    }
    if (score > bestScore) {
      secondScore = bestScore
      bestScore = score
      best = template
    } else if (score > secondScore) {
      secondScore = score
    }
  }
  if (!best || bestScore < 2 || bestScore <= secondScore) return null
  return best
}

/** One question per placeholder, in the template's own wording. */
export function fallbackQuestionsFor(placeholders: readonly string[]): string[] {
  return placeholders.map((placeholder) => `What should I put for ${placeholder}?`)
}

/**
 * Map a chat answer onto the pending placeholders. Accepted shapes:
 * one placeholder → the whole answer is its value; several → one
 * `placeholder: value` line each, or one value per line in order.
 * Returns null when the answer does not cover every placeholder, so
 * the caller asks again instead of filling with guesses.
 */
export function parsePlaceholderAnswers(
  placeholders: readonly string[],
  answerText: string,
): Map<string, string> | null {
  const text = answerText.trim()
  if (!text || placeholders.length === 0) return null
  if (placeholders.length === 1) {
    const only = placeholders[0]!
    // a lone "label: value" answer still yields just the value
    const labelled = text.match(/^(.{1,90}?)\s*[:：]\s*([\s\S]+)$/)
    if (labelled && labelled[1]!.includes(only.slice(1, -1))) {
      return new Map([[only, labelled[2]!.trim()]])
    }
    return new Map([[only, text]])
  }
  const lines = text
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
  const labelled = new Map<string, string>()
  for (const line of lines) {
    for (const placeholder of placeholders) {
      const prefix = `${placeholder}:`
      const prefixAlt = `${placeholder}：`
      if (line.startsWith(prefix) || line.startsWith(prefixAlt)) {
        labelled.set(placeholder, line.slice(prefix.length).trim())
      }
    }
  }
  if (labelled.size === placeholders.length) return labelled
  if (lines.length === placeholders.length) {
    const positional = new Map<string, string>()
    placeholders.forEach((placeholder, index) => positional.set(placeholder, lines[index]!))
    return positional
  }
  return null
}

/** The fallback's create item for a matched template + its answers. */
export function fallbackCreateItem(
  template: TemplateDef,
  answers: Map<string, string> | null,
): BuildCreateItem {
  const fill: BuildFillEntry[] = []
  if (answers) {
    for (const [find, replace] of answers) fill.push({ find, replace })
  }
  return {
    kind: 'create',
    app: template.content.app,
    title: template.title,
    templateId: template.id,
    ...(fill.length > 0 ? { fill } : {}),
  }
}

// ── The model contract ────────────────────────────────────────────────

/**
 * The system prompt for one Build chat turn. It carries the whole
 * contract: JSON only, clarify first, create-only by default, edits
 * gated, no PDF creation, and the template catalog the model may pick
 * from. Kept here (not in the service) so tests can pin its shape.
 */
export function buildChatSystemPrompt(templates: readonly TemplateDef[]): string {
  const catalog = templates
    .map(
      (template) =>
        `- ${template.id} (${template.content.app}): ${template.title} — ${template.description}`,
    )
    .join('\n')
  return [
    'You are the Build assistant inside TechPartner Office. The user describes something they need in plain words; you turn it into real files in the office apps (Docs documents, Sheets spreadsheets, Slides presentations).',
    '',
    'Reply with JSON only — no prose outside the JSON. Exactly one of these two shapes:',
    '',
    '1. When facts that would end up in a file are missing — a client or business name, amounts or line items, dates, an audience, a title — ask first, do not guess:',
    '{"questions": ["short question", "..."], "message": "one short sentence saying what you need"}',
    'Ask at most 5 questions, each about one fact. If the user has already answered a question earlier in the conversation, do not ask it again.',
    '',
    '2. When you have everything you need, give the build plan:',
    '{"plan": [item, ...], "message": "one short sentence describing what you will create"}',
    `At most ${BUILD_CHAT_MAX_ITEMS} items. One request may create files in several apps at once.`,
    '',
    'A create item is one of:',
    '- From a template: {"kind": "create", "app": "docs"|"sheets"|"slides", "title": "file title", "templateId": "<catalog id>", "fill": [{"find": "[Placeholder text]", "replace": "the user\'s fact"}]}. Prefer a template when one fits. Fill replaces bracketed placeholders in the template with facts the user gave you — never invent a name, amount or date to fill one; ask instead.',
    '- Custom document: {"kind": "create", "app": "docs", "title": "...", "markdown": "# Title\\n\\nBody in markdown"}',
    '- Custom spreadsheet: {"kind": "create", "app": "sheets", "title": "...", "sheets": [{"name": "Sheet1", "headerRow": true, "rows": [["Column", "Amount"], ["Item", 100]]}]}. A cell is a string, a number, or {"formula": "B2*C2", "cached": 500} where cached is the value the formula produces.',
    '- Custom deck: {"kind": "create", "app": "slides", "title": "...", "slides": [{"title": "...", "subtitle": "...", "bullets": ["...", "..."]}]}',
    '',
    'An edit item changes an existing document or deck by exact text replacement: {"kind": "edit", "app": "docs"|"slides", "path": "/absolute/path/file.docx", "find": "exact existing text", "replace": "new text"}. Only offer an edit when the user explicitly asks to change a specific existing file and gives its path. Edits are never combined with silent side effects, and the user will be asked to confirm the plan before any edit runs. Spreadsheets cannot be edited this way.',
    '',
    'You cannot create PDF files. If the user asks for a PDF, say so in a questions reply (ask what editable file to create instead) — never promise a PDF.',
    'Creates always make NEW files; they never overwrite anything. Titles are file titles without an extension.',
    '',
    'Template catalog (id, app, title, description):',
    catalog,
  ].join('\n')
}
