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
 * the user's facts filled in, or from custom content), edits an
 * existing file (a document or deck by exact text replacement, or a
 * spreadsheet by literal cell values), or derives a new document from
 * an existing file's content — the source is read, never changed.
 * This module validates every
 * reply before anything touches disk: an invalid reply parses to null
 * and the service asks a clarifying question instead of crashing or
 * half-building.
 *
 * SAFETY CONTRACT:
 * - Creates never overwrite: the service writes to a unique-suffixed
 *   path, and this module has no notion of an overwrite flag at all.
 * - Edits are separate item kinds (text edits for docs/slides,
 *   cell-value edits for spreadsheets — literal values only, never
 *   formulas), and the service executes them only after the user
 *   confirms in chat.
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

/** One cell assignment of a spreadsheet edit: a literal value only. */
export interface BuildCellUpdate {
  /** A1-style reference, normalized (uppercase letters, no leading zeros) */
  cell: string
  value: string | number
}

export interface BuildSheetEditItem {
  kind: 'sheet-edit'
  /** the .xlsx in the user's save folder, as a relative path */
  path: string
  /** target sheet name; omitted means the workbook's first sheet */
  sheet?: string
  updates: BuildCellUpdate[]
}

export interface BuildDeriveItem {
  kind: 'derive'
  /** the existing file to build from, as a save-folder-relative path */
  source: string
  /** what to do with its content, in the user's words */
  instruction: string
  /** title of the NEW document this creates (the source never changes) */
  title: string
}

export type BuildPlanItem = BuildCreateItem | BuildEditItem | BuildSheetEditItem | BuildDeriveItem

/** Source extensions a derive item can read (the extractors Home
 *  search uses cover exactly these editable/text formats). */
export const BUILD_DERIVE_EXTENSIONS: readonly string[] = [
  'docx',
  'xlsx',
  'pptx',
  'md',
  'html',
  'txt',
]

/** Extracted source text fed to the model for a derive, in characters.
 *  Past this the plan says plainly that only the first part was used. */
export const BUILD_DERIVE_TEXT_LIMIT = 24_000

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

/**
 * One entry of the rendered transcript — a user message, or an
 * assistant message with the structured turn behind it. The service
 * owns the transcript (the panel hydrates from it on mount) and
 * persists it across restarts, capped at the last
 * {@link BUILD_CHAT_TRANSCRIPT_LIMIT} entries. A pending edit
 * confirmation is NOT part of this state: it lives only in the
 * service's memory, so a restored confirm card can never be released
 * by a later "yes" (the service marks restored ones expired on load).
 */
export interface BuildChatTranscriptEntry {
  role: 'user' | 'assistant'
  text: string
  /** the structured service turn behind an assistant entry */
  turn?: BuildChatTurn
}

export const BUILD_CHAT_TRANSCRIPT_LIMIT = 200

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
const MAX_SHEET_UPDATES = 50
const MAX_SAVE_PATH_CHARS = 260
const MAX_INSTRUCTION_CHARS = 500

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

/**
 * A path inside the user's save folder — the only form spreadsheet
 * edits (and derive sources) may use: relative, no drive letter, no
 * absolute root, and no empty / "." / ".." segments, so resolving it
 * against the save folder can never land outside it.
 */
export function isSaveFolderRelativePath(value: string): boolean {
  const path = value.trim()
  if (!path || path.length > MAX_SAVE_PATH_CHARS) return false
  if (isAbsoluteBuildPath(path)) return false
  if (path.startsWith('/') || path.startsWith('\\')) return false
  if (/^[A-Za-z]:/.test(path)) return false
  return path
    .split(/[/\\]/)
    .every((segment) => segment !== '' && segment !== '.' && segment !== '..')
}

/**
 * Normalize an A1-style cell reference ("b12" → "B12"), or null when it
 * is not strictly one: 1–3 letters + a row number, inside the real
 * spreadsheet bounds (column XFD = 16384, row 1048576).
 */
export function normalizeCellRef(value: string): string | null {
  const match = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(value.trim())
  if (!match) return null
  const letters = match[1]!.toUpperCase()
  let column = 0
  for (const ch of letters) column = column * 26 + (ch.charCodeAt(0) - 64)
  const row = Number.parseInt(match[2]!, 10)
  if (column < 1 || column > 16384 || row < 1 || row > 1048576) return null
  return `${letters}${row}`
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

function validateSheetEditItem(entry: Record<string, unknown>): BuildSheetEditItem | null {
  const path = typeof entry.path === 'string' ? entry.path.trim() : ''
  if (!isSaveFolderRelativePath(path) || !/\.xlsx$/i.test(path)) return null
  if (
    !Array.isArray(entry.updates) ||
    entry.updates.length === 0 ||
    entry.updates.length > MAX_SHEET_UPDATES
  ) {
    return null
  }
  const item: BuildSheetEditItem = { kind: 'sheet-edit', path, updates: [] }
  if (entry.sheet !== undefined) {
    const sheet = cleanString(entry.sheet, 31)
    if (!sheet || [...sheet].some((ch) => SHEET_NAME_FORBIDDEN.includes(ch))) return null
    item.sheet = sheet
  }
  const seen = new Set<string>()
  for (const raw of entry.updates) {
    if (!isRecord(raw)) return null
    const cell = typeof raw.cell === 'string' ? normalizeCellRef(raw.cell) : null
    if (!cell || seen.has(cell)) return null
    seen.add(cell)
    const value = raw.value
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) return null
      item.updates.push({ cell, value })
    } else if (typeof value === 'string') {
      // values are literals in this action: a leading "=" would read as
      // a formula in the file, and formulas are deliberately not written
      if (value.length > MAX_CELL_CHARS || value.startsWith('=')) return null
      item.updates.push({ cell, value })
    } else {
      return null
    }
  }
  return item
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
      const item =
        raw.kind === 'edit'
          ? validateEditItem(raw)
          : raw.kind === 'sheet-edit'
            ? validateSheetEditItem(raw)
            : raw.kind === 'derive'
              ? validateDeriveItem(raw)
              : validateCreateItem(raw, templates)
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
  return items.some((item) => item.kind === 'edit' || item.kind === 'sheet-edit')
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

function describeCellValue(value: string | number): string {
  return typeof value === 'number' ? String(value) : `“${truncateMiddle(value, 40)}”`
}

function validateDeriveItem(entry: Record<string, unknown>): BuildDeriveItem | null {
  const source = typeof entry.source === 'string' ? entry.source.trim() : ''
  if (!isSaveFolderRelativePath(source)) return null
  const extension = source.split('.').pop()?.toLowerCase() ?? ''
  if (!BUILD_DERIVE_EXTENSIONS.includes(extension)) return null
  const instruction = cleanString(entry.instruction, MAX_INSTRUCTION_CHARS)
  const title = cleanString(entry.title, MAX_TITLE_CHARS)
  if (!instruction || !title) return null
  return { kind: 'derive', source, instruction, title }
}

/** One plain-language line per plan item — the plan the user sees in the
 *  chat before any file is created. */
export function describePlanItem(item: BuildPlanItem, templates: readonly TemplateDef[]): string {
  if (item.kind === 'derive') {
    return `Create a document “${item.title}” from ${baseName(item.source)}`
  }
  if (item.kind === 'edit') {
    return (
      `Edit ${baseName(item.path)}: replace “${truncateMiddle(item.find, 60)}” ` +
      `with “${truncateMiddle(item.replace, 60)}”`
    )
  }
  if (item.kind === 'sheet-edit') {
    const shown = item.updates
      .slice(0, 4)
      .map((update) => `set ${update.cell} to ${describeCellValue(update.value)}`)
    const more =
      item.updates.length > shown.length ? `, and ${item.updates.length - shown.length} more` : ''
    const sheet = item.sheet ? ` (sheet “${item.sheet}”)` : ''
    return `Edit ${baseName(item.path)}${sheet}: ${shown.join(', ')}${more}`
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

const PLACEHOLDER_PATTERN = /\[[^[\]\n]{1,80}\]/g

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
 * categories. Returns null when nothing scores — and on a tie, unless
 * exactly one of the tied templates carries [bracketed placeholders]:
 * that one asks questions before building, while the others would
 * build instantly with their sample content, and clarifying beats
 * guessing. Any other tie is still a guess, so it returns null.
 */
export function matchTemplateForMessage(
  message: string,
  templates: readonly TemplateDef[],
): TemplateDef | null {
  const messageWords = new Set(wordsOf(message))
  if (messageWords.size === 0) return null
  let bestScore = 0
  const scored: Array<{ template: TemplateDef; score: number }> = []
  for (const template of templates) {
    const titleWords = new Set(wordsOf(template.title))
    const otherWords = new Set(wordsOf(`${template.description} ${template.category}`))
    let score = 0
    for (const word of messageWords) {
      if (titleWords.has(word)) score += 2
      else if (otherWords.has(word)) score += 1
    }
    scored.push({ template, score })
    if (score > bestScore) bestScore = score
  }
  if (bestScore < 2) return null
  const tied = scored.filter((entry) => entry.score === bestScore).map((entry) => entry.template)
  if (tied.length === 1) return tied[0]!
  const asking = tied.filter((template) => extractTemplatePlaceholders(template).length > 0)
  return asking.length === 1 ? asking[0]! : null
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
    'An edit item changes an existing document or deck by exact text replacement: {"kind": "edit", "app": "docs"|"slides", "path": "/absolute/path/file.docx", "find": "exact existing text", "replace": "new text"}. Only offer an edit when the user explicitly asks to change a specific existing file and gives its path.',
    'A spreadsheet edit item sets cell values in an existing spreadsheet: {"kind": "sheet-edit", "path": "Budget.xlsx", "updates": [{"cell": "B12", "value": 4500}, {"cell": "C3", "value": "Total"}]}. The path is relative to the user\'s save folder and must end in .xlsx. Add "sheet": "Sheet name" only when the user names the sheet — otherwise the first sheet is edited. Values are literal strings or numbers: formulas are not supported, so never start a value with "=". At most 50 cells in one item. Only offer it when the user names the spreadsheet and the cells to change.',
    'Edits of either kind are never combined with silent side effects, and the user will be asked to confirm the plan before any edit runs.',
    'A derive item creates a NEW document based on an existing file: {"kind": "derive", "source": "Budget.xlsx", "instruction": "summarize it as a one-page brief", "title": "Budget Summary"}. The source path is relative to the save folder; supported sources are .docx, .xlsx, .pptx, .md, .html and .txt. The source file is never changed — the output is always a new document. Use it when the user asks to summarize, rewrite, reformat or build on a file they already have, and put their exact ask in the instruction.',
    '',
    'You cannot create PDF files. If the user asks for a PDF, say so in a questions reply (ask what editable file to create instead) — never promise a PDF.',
    'Creates always make NEW files; they never overwrite anything. Titles are file titles without an extension.',
    '',
    'Template catalog (id, app, title, description):',
    catalog,
  ].join('\n')
}

/**
 * The system prompt for the second, content-writing call of a derive
 * item: the service hands the model one source file's extracted text
 * plus the user's instruction, and gets back the new document as
 * markdown. Kept separate from the plan contract on purpose — this
 * call writes prose, never JSON, and may only use the source's facts.
 */
export function deriveDocumentSystemPrompt(): string {
  return [
    "You are the writing assistant inside TechPartner Office. You are given the text of one of the user's existing files and an instruction for a new document based on it.",
    'Reply with the new document in markdown only — no commentary before or after it, no code fences around it. Start with a top-level heading.',
    'Use only facts from the source text. Do not invent names, figures, dates or claims that are not in it; when the source does not cover something the instruction asks for, leave it out rather than guessing.',
  ].join('\n')
}
