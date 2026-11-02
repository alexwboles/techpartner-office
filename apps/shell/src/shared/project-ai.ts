/**
 * Project AI (pilot: business plan) — pure, framework-free core.
 *
 * This module runs in the shell main process, the shell renderer, and
 * vitest. It must not import node or Electron APIs.
 *
 * SAFETY CONTRACT (the whole pilot rests on it):
 * - Only two suggestion classes may ever auto-apply: typo fixes and
 *   number syncs where a figure in the plan document or pitch deck
 *   disagrees with the budget spreadsheet. The spreadsheet is the
 *   source of truth; the document/deck figure is corrected to it.
 * - Everything else is an approval-queue suggestion and applies only
 *   when the user clicks accept.
 * - A number sync is SAFE only when exactly one budget row matches the
 *   mention. Ambiguous matches (zero rows is no suggestion; two or more
 *   rows is a queue item, never an auto-apply) are never auto-applied.
 * - Suggestions produced by the language model are always queue items:
 *   `autoSafe` is true only for deterministic typo/number-sync findings
 *   built by `reviewProjectContent` below.
 */

export const PROJECT_AI_DOC = 'business-plan.docx'
export const PROJECT_AI_SHEET = 'business-plan-budget.xlsx'
export const PROJECT_AI_DECK = 'business-plan-pitch.pptx'

export type ProjectAiFileKind = 'doc' | 'sheet' | 'deck'

export type ProjectAiSuggestionKind = 'typo' | 'number-sync' | 'other'

export interface ProjectAiSuggestion {
  id: string
  kind: ProjectAiSuggestionKind
  /** which project file the suggestion edits ('sheet' is never auto-edited) */
  file: ProjectAiFileKind
  fileName: string
  title: string
  detail: string
  /** exact text to find in the file; empty for free-form queue items */
  before: string
  /** exact replacement text */
  after: string
  /** true only for verified typo / number-sync findings (see contract) */
  autoSafe: boolean
  /** set when the finding was ambiguous and therefore queued, not applied */
  ambiguity?: string
}

export interface ProjectAiChangeLogEntry {
  at: string
  suggestionId: string
  kind: ProjectAiSuggestionKind
  fileName: string
  before: string
  after: string
  /** how the change was applied: 'auto' for SAFE, 'accepted' for queue */
  via: 'auto' | 'accepted'
}

export interface ProjectAiState {
  folder: string | null
  files: string[]
  status: 'idle' | 'building' | 'ready' | 'error'
  statusMessage: string
  /** SAFE suggestions that were auto-applied (kept for the change log view) */
  applied: ProjectAiChangeLogEntry[]
  /** suggestions waiting on the user */
  queue: ProjectAiSuggestion[]
  lastReviewAt: string | null
  lastError: string | null
}

export function projectAiFileName(kind: ProjectAiFileKind): string {
  return kind === 'doc' ? PROJECT_AI_DOC : kind === 'sheet' ? PROJECT_AI_SHEET : PROJECT_AI_DECK
}

/** The two classes that may auto-apply. Everything else is queue-only. */
export function isSafeSuggestionClass(kind: ProjectAiSuggestionKind): boolean {
  return kind === 'typo' || kind === 'number-sync'
}

/**
 * The single gate the service checks before any auto-apply: the class must
 * be SAFE *and* this finding must have been verified (exactly one source
 * row, exact text present). A suggestion that fails either check is queued.
 */
export function mayAutoApply(suggestion: ProjectAiSuggestion): boolean {
  return suggestion.autoSafe && isSafeSuggestionClass(suggestion.kind) && suggestion.before !== ''
}

export function formatMoney(amount: number): string {
  const rounded = Math.round(amount * 100) / 100
  const text = Number.isInteger(rounded)
    ? rounded.toLocaleString('en-US')
    : rounded.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  return `$${text}`
}

export function parseMoney(raw: string): number | null {
  const cleaned = raw.replace(/[$,\s]/g, '')
  if (!/^\d+(\.\d{1,2})?$/.test(cleaned)) return null
  const value = Number(cleaned)
  return Number.isFinite(value) ? value : null
}

function normalizeWords(text: string): string[] {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean)
}

function normalizeLabel(label: string): string {
  return normalizeWords(label).join(' ')
}

export interface BudgetRow {
  label: string
  amount: number
}

/** Unambiguous, whole-word typos. The map is deliberately small: a word
 *  that could be intentional (a name, a term of art) must never be in it. */
export const TYPO_FIXES: Record<string, string> = {
  teh: 'the',
  recieve: 'receive',
  seperate: 'separate',
  occured: 'occurred',
  buisness: 'business',
  bussiness: 'business',
  planing: 'planning',
  managment: 'management',
  finacial: 'financial',
  marketting: 'marketing',
  competiton: 'competition',
  enviroment: 'environment',
  goverment: 'government',
  neccessary: 'necessary',
  accomodate: 'accommodate',
  adress: 'address',
  calender: 'calendar',
  definately: 'definitely',
  experiance: 'experience',
  independant: 'independent',
  noticable: 'noticeable',
  posession: 'possession',
  succesful: 'successful',
  untill: 'until',
  withing: 'within',
  ammount: 'amount',
  budjet: 'budget',
  costumers: 'customers',
  revenu: 'revenue',
  expences: 'expenses',
  liablities: 'liabilities',
  assests: 'assets',
  opertations: 'operations',
  stratagy: 'strategy',
  growht: 'growth',
  fundding: 'funding',
  investers: 'investors',
  partenership: 'partnership',
  liscense: 'license',
}

function preserveCase(original: string, replacement: string): string {
  if (original[0] === original[0]?.toUpperCase()) {
    return replacement[0]?.toUpperCase() + replacement.slice(1)
  }
  return replacement
}

/** Whole-word typo findings in `text` (doc/deck only — spreadsheet cells
 *  are labels and formulas, so sheet typos are never auto-fixed here). */
export function findTypos(
  text: string,
  file: ProjectAiFileKind,
  startIndex: number,
): ProjectAiSuggestion[] {
  const out: ProjectAiSuggestion[] = []
  let index = startIndex
  const pattern = /\b([A-Za-z]+)\b/g
  let match: RegExpExecArray | null
  while ((match = pattern.exec(text)) !== null) {
    const word = match[1]
    const fix = TYPO_FIXES[word.toLowerCase()]
    if (!fix) continue
    index += 1
    out.push({
      id: `typo-${file}-${index}`,
      kind: 'typo',
      file,
      fileName: projectAiFileName(file),
      title: `Typo: ${word} → ${preserveCase(word, fix)}`,
      detail: `Fixed a spelling error in ${projectAiFileName(file)}.`,
      before: word,
      after: preserveCase(word, fix),
      autoSafe: true,
    })
  }
  return out
}

export interface MoneyMention {
  /** raw matched text, e.g. "$8,500" */
  raw: string
  amount: number
  /** words immediately before the amount, used to identify the figure */
  context: string
}

const MONEY_PATTERN = /\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?|\d+(?:\.\d{2})?)/g

export function extractMoneyMentions(text: string): MoneyMention[] {
  const out: MoneyMention[] = []
  let match: RegExpExecArray | null
  while ((match = MONEY_PATTERN.exec(text)) !== null) {
    const amount = parseMoney(match[0])
    if (amount === null) continue
    const before = text.slice(Math.max(0, match.index - 80), match.index)
    out.push({ raw: match[0], amount, context: before })
  }
  return out
}

function rowMatchesContext(row: BudgetRow, context: string): boolean {
  const labelWords = normalizeWords(row.label)
  if (labelWords.length === 0) return false
  const contextWords = new Set(normalizeWords(context))
  return labelWords.every((word) => contextWords.has(word))
}

/**
 * Number-sync findings for one text file. A mention is SAFE only when
 * exactly one budget row matches its context and holds a different amount.
 * Mentions matching several rows are queued with an ambiguity note; the
 * spreadsheet is never edited by a sync (it is the source of truth).
 */
export function findNumberSyncs(
  text: string,
  file: Exclude<ProjectAiFileKind, 'sheet'>,
  budgetRows: BudgetRow[],
  startIndex: number,
): ProjectAiSuggestion[] {
  const out: ProjectAiSuggestion[] = []
  let index = startIndex
  for (const mention of extractMoneyMentions(text)) {
    const candidates = budgetRows.filter((row) => rowMatchesContext(row, mention.context))
    const differing = candidates.filter((row) => Math.abs(row.amount - mention.amount) > 0.005)
    if (differing.length === 0) continue
    index += 1
    if (differing.length === 1) {
      const row = differing[0]
      out.push({
        id: `number-sync-${file}-${index}`,
        kind: 'number-sync',
        file,
        fileName: projectAiFileName(file),
        title: `${row.label}: ${mention.raw} → ${formatMoney(row.amount)}`,
        detail:
          `The budget spreadsheet lists ${row.label} as ${formatMoney(row.amount)}, ` +
          `but ${projectAiFileName(file)} says ${mention.raw}. The spreadsheet is the source of truth, so the figure was corrected to it.`,
        before: mention.raw,
        after: formatMoney(row.amount),
        autoSafe: true,
      })
    } else {
      out.push({
        id: `number-sync-${file}-${index}`,
        kind: 'number-sync',
        file,
        fileName: projectAiFileName(file),
        title: `${mention.raw} disagrees with the budget (ambiguous)`,
        detail:
          `This figure could match several budget rows (${differing.map((row) => `${row.label} ${formatMoney(row.amount)}`).join(', ')}), ` +
          `so it was not changed automatically. Accept the correction you mean, or leave it.`,
        before: mention.raw,
        after: mention.raw,
        autoSafe: false,
        ambiguity: `matches ${differing.length} budget rows`,
      })
    }
  }
  return out
}

export interface ProjectContent {
  docText: string
  deckText: string
  budgetRows: BudgetRow[]
}

/**
 * Deterministic cross-file review. Returns SAFE findings (auto-applied by
 * the service) and queue findings (shown for approval). The language model
 * never contributes to `safe`: anything it suggests is appended to the
 * queue by the service as kind 'other'.
 */
export function reviewProjectContent(content: ProjectContent): {
  safe: ProjectAiSuggestion[]
  queued: ProjectAiSuggestion[]
} {
  const safe: ProjectAiSuggestion[] = []
  const queued: ProjectAiSuggestion[] = []
  let typoIndex = 0
  let syncIndex = 0
  const collect = (suggestions: ProjectAiSuggestion[]): void => {
    for (const suggestion of suggestions) {
      if (mayAutoApply(suggestion)) safe.push(suggestion)
      else queued.push(suggestion)
    }
  }
  const docTypos = findTypos(content.docText, 'doc', typoIndex)
  typoIndex += docTypos.length
  const deckTypos = findTypos(content.deckText, 'deck', typoIndex)
  collect(docTypos)
  collect(deckTypos)
  const docSyncs = findNumberSyncs(content.docText, 'doc', content.budgetRows, syncIndex)
  syncIndex += docSyncs.length
  const deckSyncs = findNumberSyncs(content.deckText, 'deck', content.budgetRows, syncIndex)
  collect(docSyncs)
  collect(deckSyncs)
  return { safe, queued }
}

// ── Business-plan content (template; the language model may replace it) ──

export interface BusinessPlanDraft {
  businessName: string
  sections: Array<{ heading: string; body: string }>
  budget: BudgetRow[]
}

export const DEFAULT_BUDGET: BudgetRow[] = [
  { label: 'Equipment', amount: 8500 },
  { label: 'Initial inventory', amount: 6000 },
  { label: 'Marketing launch', amount: 3500 },
  { label: 'Licenses and permits', amount: 1200 },
  { label: 'Website and branding', amount: 2500 },
  { label: 'Rent deposit', amount: 3000 },
  { label: 'Working capital reserve', amount: 10000 },
]

export function budgetTotal(budget: BudgetRow[]): number {
  return budget.reduce((sum, row) => sum + row.amount, 0)
}

export function businessNameFromBrief(brief: string): string {
  const first = brief.trim().split(/[.\n]/)[0] ?? ''
  const cleaned = first.replace(/^a\s+/i, '').trim()
  if (!cleaned) return 'New Business'
  return cleaned.length > 80 ? `${cleaned.slice(0, 77).trim()}…` : cleaned
}

/** Deterministic plan content. Used whenever the model is unavailable,
 *  and as the shape the model's draft is validated against. */
export function buildBusinessPlanContent(brief: string): BusinessPlanDraft {
  const name = businessNameFromBrief(brief)
  const total = formatMoney(budgetTotal(DEFAULT_BUDGET))
  const rows = DEFAULT_BUDGET.map((row) => `- ${row.label}: ${formatMoney(row.amount)}`).join('\n')
  return {
    businessName: name,
    sections: [
      {
        heading: 'Executive summary',
        body:
          `${name}. ${brief.trim()}\n\n` +
          `This plan sets out how the business will launch, reach its first customers, and become self-sustaining. ` +
          `Total startup costs are ${total}.`,
      },
      {
        heading: 'The business',
        body:
          `The business offers a focused product or service to a local market first, then grows by reputation and repeat customers. ` +
          `Operations stay lean: a small team, simple systems, and costs tracked against the budget in this project.`,
      },
      {
        heading: 'Market',
        body:
          `Customers are local buyers who value quality, convenience, and personal service. ` +
          `The plan competes on those strengths rather than on price alone, and marketing concentrates on the channels where those customers already look.`,
      },
      {
        heading: 'Financial summary',
        body:
          `Startup budget:\n${rows}\n\nTotal startup costs: ${total}. ` +
          `The budget spreadsheet in this project is the source of truth for every figure in this plan and in the pitch deck.`,
      },
    ],
    budget: [...DEFAULT_BUDGET],
  }
}

/** Markdown for the plan document. */
export function businessPlanMarkdown(draft: BusinessPlanDraft): string {
  const lines = [`# Business plan: ${draft.businessName}`, '']
  for (const section of draft.sections) {
    lines.push(`## ${section.heading}`, '', section.body, '')
  }
  return lines.join('\n')
}

/** Row matrix for the budget spreadsheet (header, rows, total formula). */
export function businessPlanBudgetRows(draft: BusinessPlanDraft): (string | number)[][] {
  const rows: (string | number)[][] = [['Category', 'Amount']]
  draft.budget.forEach((row) => rows.push([row.label, row.amount]))
  rows.push(['Total', `=SUM(B2:B${draft.budget.length + 1})`])
  return rows
}

/** Markdown outline for the pitch deck (one `#` title per slide). */
export function businessPlanDeckOutline(draft: BusinessPlanDraft): string {
  const total = formatMoney(budgetTotal(draft.budget))
  return [
    `# ${draft.businessName}`,
    '- A business plan, a budget, and a pitch — in one project.',
    '',
    '# The opportunity',
    '- Local customers who value quality, convenience, and personal service.',
    '- A focused offer, built to grow by reputation and repeat business.',
    '',
    '# What we offer',
    '- A simple, well-run operation with a small team and lean systems.',
    '- Costs tracked against the budget in this project.',
    '',
    '# Financials',
    `- Total startup costs: ${total}.`,
    ...draft.budget.map((row) => `- ${row.label}: ${formatMoney(row.amount)}.`),
    '',
    '# Next steps',
    '- Confirm the launch budget and the first marketing push.',
    '- Open the plan document and the budget to review the details.',
  ].join('\n')
}
