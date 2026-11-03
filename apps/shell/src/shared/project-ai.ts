/**
 * Project AI — pure, framework-free core.
 *
 * Project types (business plan, grant application, event plan and
 * marketing plan) are data, not separate flows: one per-type definition
 * (file names, document sections, budget rows, deck slides, prompt
 * framing) is consumed by the single build / review flow.
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
export const PROJECT_AI_GRANT_DOC = 'grant-application.docx'
export const PROJECT_AI_GRANT_SHEET = 'grant-application-budget.xlsx'
export const PROJECT_AI_GRANT_DECK = 'grant-application-overview.pptx'
export const PROJECT_AI_EVENT_DOC = 'event-plan.docx'
export const PROJECT_AI_EVENT_SHEET = 'event-plan-budget.xlsx'
export const PROJECT_AI_EVENT_DECK = 'event-plan-overview.pptx'
export const PROJECT_AI_MARKETING_DOC = 'marketing-plan.docx'
export const PROJECT_AI_MARKETING_SHEET = 'marketing-plan-budget.xlsx'
export const PROJECT_AI_MARKETING_DECK = 'marketing-plan-overview.pptx'

export type ProjectAiFileKind = 'doc' | 'sheet' | 'deck'

/** The project types Project AI can build. Business plan is the default:
 *  state saved before project types existed carries no type and must
 *  read as a business plan. */
export type ProjectAiProjectType =
  'business-plan' | 'grant-application' | 'event-plan' | 'marketing-plan'

export const PROJECT_AI_PROJECT_TYPES: readonly ProjectAiProjectType[] = [
  'business-plan',
  'grant-application',
  'event-plan',
  'marketing-plan',
]

export const DEFAULT_PROJECT_AI_TYPE: ProjectAiProjectType = 'business-plan'

export function isProjectAiProjectType(value: unknown): value is ProjectAiProjectType {
  return (
    typeof value === 'string' && (PROJECT_AI_PROJECT_TYPES as readonly string[]).includes(value)
  )
}

/** Resolve any stored / incoming type value; missing or unknown values
 *  fall back to the business plan so old saved state keeps working. */
export function normalizeProjectType(value: unknown): ProjectAiProjectType {
  return isProjectAiProjectType(value) ? value : DEFAULT_PROJECT_AI_TYPE
}

const PROJECT_AI_FILE_NAMES: Record<ProjectAiProjectType, Record<ProjectAiFileKind, string>> = {
  'business-plan': { doc: PROJECT_AI_DOC, sheet: PROJECT_AI_SHEET, deck: PROJECT_AI_DECK },
  'grant-application': {
    doc: PROJECT_AI_GRANT_DOC,
    sheet: PROJECT_AI_GRANT_SHEET,
    deck: PROJECT_AI_GRANT_DECK,
  },
  'event-plan': {
    doc: PROJECT_AI_EVENT_DOC,
    sheet: PROJECT_AI_EVENT_SHEET,
    deck: PROJECT_AI_EVENT_DECK,
  },
  'marketing-plan': {
    doc: PROJECT_AI_MARKETING_DOC,
    sheet: PROJECT_AI_MARKETING_SHEET,
    deck: PROJECT_AI_MARKETING_DECK,
  },
}

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
  /** the project type being built / watched; absent in old saved state */
  projectType: ProjectAiProjectType
  status: 'idle' | 'building' | 'ready' | 'error'
  statusMessage: string
  /** SAFE suggestions that were auto-applied (kept for the change log view) */
  applied: ProjectAiChangeLogEntry[]
  /** suggestions waiting on the user */
  queue: ProjectAiSuggestion[]
  lastReviewAt: string | null
  lastError: string | null
}

/** The type a state belongs to; state saved before project types existed
 *  has no type and reads as a business plan. */
export function projectTypeOfState(
  state: { projectType?: unknown } | null | undefined,
): ProjectAiProjectType {
  return normalizeProjectType(state?.projectType)
}

/** Fill a possibly-old partial state out to a complete one, defaulting a
 *  missing project type to the business plan. */
export function normalizeProjectAiState(
  state: Partial<ProjectAiState> | null | undefined,
): ProjectAiState {
  return {
    folder: state?.folder ?? null,
    files: state?.files ? [...state.files] : [],
    projectType: normalizeProjectType(state?.projectType),
    status: state?.status ?? 'idle',
    statusMessage: state?.statusMessage ?? '',
    applied: state?.applied ? [...state.applied] : [],
    queue: state?.queue ? state.queue.map((suggestion) => ({ ...suggestion })) : [],
    lastReviewAt: state?.lastReviewAt ?? null,
    lastError: state?.lastError ?? null,
  }
}

export function projectAiFileName(
  kind: ProjectAiFileKind,
  projectType: ProjectAiProjectType = DEFAULT_PROJECT_AI_TYPE,
): string {
  return PROJECT_AI_FILE_NAMES[normalizeProjectType(projectType)][kind]
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
  projectType: ProjectAiProjectType = DEFAULT_PROJECT_AI_TYPE,
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
      fileName: projectAiFileName(file, projectType),
      title: `Typo: ${word} → ${preserveCase(word, fix)}`,
      detail: `Fixed a spelling error in ${projectAiFileName(file, projectType)}.`,
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
  projectType: ProjectAiProjectType = DEFAULT_PROJECT_AI_TYPE,
): ProjectAiSuggestion[] {
  const out: ProjectAiSuggestion[] = []
  let index = startIndex
  for (const mention of extractMoneyMentions(text)) {
    const candidates = budgetRows.filter((row) => rowMatchesContext(row, mention.context))
    // If ANY matching row already holds this amount, the figure is
    // consistent with the budget as written — correcting it to a
    // different matching row would invent a discrepancy, not fix one.
    if (candidates.some((row) => Math.abs(row.amount - mention.amount) <= 0.005)) continue
    if (candidates.length === 0) continue
    index += 1
    if (candidates.length === 1) {
      const row = candidates[0]
      out.push({
        id: `number-sync-${file}-${index}`,
        kind: 'number-sync',
        file,
        fileName: projectAiFileName(file, projectType),
        title: `${row.label}: ${mention.raw} → ${formatMoney(row.amount)}`,
        detail:
          `The budget spreadsheet lists ${row.label} as ${formatMoney(row.amount)}, ` +
          `but ${projectAiFileName(file, projectType)} says ${mention.raw}. The spreadsheet is the source of truth, so the figure was corrected to it.`,
        before: mention.raw,
        after: formatMoney(row.amount),
        autoSafe: true,
      })
    } else {
      out.push({
        id: `number-sync-${file}-${index}`,
        kind: 'number-sync',
        file,
        fileName: projectAiFileName(file, projectType),
        title: `${mention.raw} disagrees with the budget (ambiguous)`,
        detail:
          `This figure could match several budget rows (${candidates.map((row) => `${row.label} ${formatMoney(row.amount)}`).join(', ')}), ` +
          `so it was not changed automatically. Accept the correction you mean, or leave it.`,
        before: mention.raw,
        after: mention.raw,
        autoSafe: false,
        ambiguity: `matches ${candidates.length} budget rows`,
      })
    }
  }
  return out
}

export interface ProjectContent {
  docText: string
  deckText: string
  budgetRows: BudgetRow[]
  /** which project type the files belong to (defaults to business plan) */
  projectType?: ProjectAiProjectType
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
  const projectType = normalizeProjectType(content.projectType)
  const docTypos = findTypos(content.docText, 'doc', typoIndex, projectType)
  typoIndex += docTypos.length
  const deckTypos = findTypos(content.deckText, 'deck', typoIndex, projectType)
  collect(docTypos)
  collect(deckTypos)
  const docSyncs = findNumberSyncs(
    content.docText,
    'doc',
    content.budgetRows,
    syncIndex,
    projectType,
  )
  syncIndex += docSyncs.length
  const deckSyncs = findNumberSyncs(
    content.deckText,
    'deck',
    content.budgetRows,
    syncIndex,
    projectType,
  )
  collect(docSyncs)
  collect(deckSyncs)
  return { safe, queued }
}

/**
 * Model suggestions, validated. Model output is ALWAYS a queue item —
 * kind 'other', autoSafe false — whatever the project type and however
 * much an entry looks like a typo or a number fix: only deterministic
 * findings from reviewProjectContent may ever auto-apply.
 */
export function validateModelSuggestions(
  value: unknown,
  startIndex: number,
  projectType: unknown = DEFAULT_PROJECT_AI_TYPE,
): ProjectAiSuggestion[] {
  if (!Array.isArray(value)) return []
  const type = normalizeProjectType(projectType)
  const out: ProjectAiSuggestion[] = []
  let index = startIndex
  for (const entry of value.slice(0, 20)) {
    if (!isDraftRecord(entry)) continue
    const file = entry.file === 'doc' || entry.file === 'deck' ? entry.file : null
    const title = typeof entry.title === 'string' ? entry.title.trim() : ''
    const detail = typeof entry.detail === 'string' ? entry.detail.trim() : ''
    const before = typeof entry.before === 'string' ? entry.before : ''
    const after = typeof entry.after === 'string' ? entry.after : ''
    if (!file || !title || !before || !after || before === after) continue
    index += 1
    out.push({
      id: `other-${file}-${index}`,
      kind: 'other',
      file,
      fileName: projectAiFileName(file, type),
      title,
      detail,
      before,
      after,
      autoSafe: false,
    })
  }
  return out
}

// ── Project content (per-type templates; the language model may replace it) ──

/**
 * A draft of any project type. `businessName` is retained as the display
 * title for compatibility with the original business-plan shape; `title`
 * is the same value under a type-neutral name.
 */
export interface ProjectDraft {
  projectType: ProjectAiProjectType
  title: string
  businessName: string
  sections: Array<{ heading: string; body: string }>
  budget: BudgetRow[]
}

export type BusinessPlanDraft = ProjectDraft

export const DEFAULT_BUDGET: BudgetRow[] = [
  { label: 'Equipment', amount: 8500 },
  { label: 'Initial inventory', amount: 6000 },
  { label: 'Marketing launch', amount: 3500 },
  { label: 'Licenses and permits', amount: 1200 },
  { label: 'Website and branding', amount: 2500 },
  { label: 'Rent deposit', amount: 3000 },
  { label: 'Working capital reserve', amount: 10000 },
]

export const GRANT_DEFAULT_BUDGET: BudgetRow[] = [
  { label: 'Personnel', amount: 18000 },
  { label: 'Supplies', amount: 4500 },
  { label: 'Equipment', amount: 7500 },
  { label: 'Other costs', amount: 3000 },
]

export const EVENT_DEFAULT_BUDGET: BudgetRow[] = [
  { label: 'Venue', amount: 5200 },
  { label: 'Catering', amount: 6800 },
  { label: 'Materials', amount: 1800 },
  { label: 'Promotion', amount: 2400 },
]

export const MARKETING_DEFAULT_BUDGET: BudgetRow[] = [
  { label: 'Social media - Month 1', amount: 1500 },
  { label: 'Search advertising - Month 1', amount: 2000 },
  { label: 'Content creation - Month 1', amount: 1200 },
  { label: 'Social media - Month 2', amount: 1800 },
  { label: 'Search advertising - Month 2', amount: 2200 },
  { label: 'Email marketing - Month 2', amount: 900 },
]

export function budgetTotal(budget: BudgetRow[]): number {
  return budget.reduce((sum, row) => sum + row.amount, 0)
}

export interface ProjectAiSectionTemplate {
  heading: string
  /** body text; may use {name}, {brief}, {total} and {budgetLines} */
  body: string
}

export interface ProjectAiSlideTemplate {
  /** slide title; may use {name} and {total} */
  heading: string
  /** bullets; may use {name} / {total}, and a lone {budget} bullet expands
   *  to one bullet per budget row */
  bullets: string[]
}

/**
 * The per-type definition the single build flow consumes. Adding a type
 * is a new entry here plus its file names above — the service, the
 * review, and the safety gate do not change per type.
 */
export interface ProjectAiTypeDefinition {
  id: ProjectAiProjectType
  /** document title prefix, e.g. "Business plan" in "# Business plan: X" */
  docTitle: string
  /** name of the budget sheet tab */
  sheetName: string
  /** phrase for the budget total, e.g. "Total requested" */
  totalLabel: string
  /** model prompt framing: what the model drafts, e.g. "grant applications" */
  draftSubject: string
  /** model prompt framing: what the model reviews, e.g. "grant application project" */
  reviewSubject: string
  /** label introducing the brief in the model draft prompt */
  briefPromptLabel: string
  /** build status summary template; may use {total} */
  buildSummary: string
  defaultBudget: BudgetRow[]
  sections: ProjectAiSectionTemplate[]
  slides: ProjectAiSlideTemplate[]
}

export const PROJECT_AI_TYPE_DEFINITIONS: Record<ProjectAiProjectType, ProjectAiTypeDefinition> = {
  'business-plan': {
    id: 'business-plan',
    docTitle: 'Business plan',
    sheetName: 'Budget',
    totalLabel: 'Total startup costs',
    draftSubject: 'small business plans',
    reviewSubject: 'business plan project',
    briefPromptLabel: 'Business idea / brief:',
    buildSummary: 'plan, budget ({total} total) and pitch deck',
    defaultBudget: DEFAULT_BUDGET,
    sections: [
      {
        heading: 'Executive summary',
        body:
          '{name}. {brief}\n\n' +
          'This plan sets out how the business will launch, reach its first customers, and become self-sustaining. ' +
          'Total startup costs are {total}.',
      },
      {
        heading: 'The business',
        body:
          'The business offers a focused product or service to a local market first, then grows by reputation and repeat customers. ' +
          'Operations stay lean: a small team, simple systems, and costs tracked against the budget in this project.',
      },
      {
        heading: 'Market',
        body:
          'Customers are local buyers who value quality, convenience, and personal service. ' +
          'The plan competes on those strengths rather than on price alone, and marketing concentrates on the channels where those customers already look.',
      },
      {
        heading: 'Financial summary',
        body:
          'Startup budget:\n{budgetLines}\n\nTotal startup costs: {total}. ' +
          'The budget spreadsheet in this project is the source of truth for every figure in this plan and in the pitch deck.',
      },
    ],
    slides: [
      { heading: '{name}', bullets: ['A business plan, a budget, and a pitch — in one project.'] },
      {
        heading: 'The opportunity',
        bullets: [
          'Local customers who value quality, convenience, and personal service.',
          'A focused offer, built to grow by reputation and repeat business.',
        ],
      },
      {
        heading: 'What we offer',
        bullets: [
          'A simple, well-run operation with a small team and lean systems.',
          'Costs tracked against the budget in this project.',
        ],
      },
      { heading: 'Financials', bullets: ['Total startup costs: {total}.', '{budget}'] },
      {
        heading: 'Next steps',
        bullets: [
          'Confirm the launch budget and the first marketing push.',
          'Open the plan document and the budget to review the details.',
        ],
      },
    ],
  },
  'grant-application': {
    id: 'grant-application',
    docTitle: 'Grant application',
    sheetName: 'Grant budget',
    totalLabel: 'Total requested',
    draftSubject: 'grant applications',
    reviewSubject: 'grant application project',
    briefPromptLabel: 'Grant project brief:',
    buildSummary: 'application, budget ({total} total requested) and overview deck',
    defaultBudget: GRANT_DEFAULT_BUDGET,
    sections: [
      {
        heading: 'Need statement',
        body:
          '{brief}\n\n' +
          'The need this project addresses is specific and current: people in the community are affected now, and existing resources do not fully meet the need. ' +
          'This section states who is affected and what changes if the work is funded.',
      },
      {
        heading: 'Project description',
        body:
          'The project will deliver the work described in the brief, on a clear timeline and with named responsibility for each part. ' +
          'Activities follow the budget in this project: personnel lead delivery, supplies and equipment support it, and other costs cover the remainder.',
      },
      {
        heading: 'Goals and outcomes',
        body:
          'Goals: deliver the project as described, reach the people named in the need statement, and finish on budget. ' +
          'Outcomes are the measurable changes the project produces — what will be different, for whom, and by when — recorded as the work proceeds.',
      },
      {
        heading: 'Evaluation',
        body:
          'Progress is checked against the goals and outcomes above at set points during the project. ' +
          'Spending is tracked against the budget spreadsheet in this project, which is the source of truth for every figure in this application and in the overview deck.',
      },
      {
        heading: 'Organizational capacity',
        body:
          'The organization has the people, experience, and systems to carry out this project and account for the funding. ' +
          'Roles are assigned before work starts, records are kept throughout, and a final report closes the project.',
      },
      {
        heading: 'Budget summary',
        body:
          'Grant budget:\n{budgetLines}\n\nTotal requested: {total}. ' +
          'The budget spreadsheet in this project is the source of truth for every figure in this application.',
      },
    ],
    slides: [
      {
        heading: '{name}',
        bullets: ['A grant application, a budget, and an overview — in one project.'],
      },
      {
        heading: 'Project overview',
        bullets: [
          'The need, the project, and the change it will make.',
          'Described in full in the application document in this project.',
        ],
      },
      {
        heading: 'Need and goals',
        bullets: [
          'A specific, current need in the community.',
          'Clear goals and measurable outcomes, checked as the work proceeds.',
        ],
      },
      { heading: 'Requested funding', bullets: ['Total requested: {total}.', '{budget}'] },
      {
        heading: 'Next steps',
        bullets: [
          'Confirm the requested amount and the submission deadline.',
          'Open the application document and the budget to review the details.',
        ],
      },
    ],
  },
  'event-plan': {
    id: 'event-plan',
    docTitle: 'Event plan',
    sheetName: 'Event budget',
    totalLabel: 'Total event budget',
    draftSubject: 'event plans',
    reviewSubject: 'event plan project',
    briefPromptLabel: 'Event brief:',
    buildSummary: 'plan, budget ({total} total) and overview deck',
    defaultBudget: EVENT_DEFAULT_BUDGET,
    sections: [
      {
        heading: 'Overview',
        body:
          '{brief}\n\n' +
          'This plan brings the event together in one place: what will happen, when, where, and who is responsible for each part. ' +
          'Total event budget is {total}.',
      },
      {
        heading: 'Schedule / run of show',
        body:
          'Run of show, in order: doors and setup, welcome, the main programme, breaks as needed, close, and teardown. ' +
          'Times are confirmed with the venue and shared with everyone in a role before the day.',
      },
      {
        heading: 'Venue & logistics',
        body:
          'The venue, access, equipment, signage, and supplies are confirmed in advance and checked on the day. ' +
          'Catering numbers follow the guest count, and a named contact owns each delivery.',
      },
      {
        heading: 'Roles',
        body:
          'Every part of the event has a named owner: host, setup, welcome, programme, catering contact, and teardown. ' +
          'No task is left unowned.',
      },
      {
        heading: 'Checklist',
        body:
          'Before the day: book the venue, confirm catering, order materials, and send promotion. ' +
          'On the day: set up, run the schedule, and close down. After: thank attendees and sponsors and record final costs.',
      },
      {
        heading: 'Budget summary',
        body:
          'Event budget:\n{budgetLines}\n\nTotal event budget: {total}. ' +
          'The budget spreadsheet in this project is the source of truth for every figure in this plan.',
      },
    ],
    slides: [
      {
        heading: '{name}',
        bullets: ['An event plan, a budget, and an overview — in one project.'],
      },
      {
        heading: 'Event overview',
        bullets: [
          'What will happen, when, and where — in one view.',
          'Schedule, venue, and roles are set out in the planning document.',
        ],
      },
      {
        heading: 'Schedule & logistics',
        bullets: [
          'A run of show from setup to teardown.',
          'Venue, catering, and materials confirmed before the day.',
        ],
      },
      { heading: 'Event budget', bullets: ['Total event budget: {total}.', '{budget}'] },
      {
        heading: 'Next steps',
        bullets: [
          'Confirm the venue and the date.',
          'Open the planning document and the budget to review the details.',
        ],
      },
    ],
  },
  'marketing-plan': {
    id: 'marketing-plan',
    docTitle: 'Marketing plan',
    sheetName: 'Campaign budget',
    totalLabel: 'Total campaign budget',
    draftSubject: 'marketing plans',
    reviewSubject: 'marketing plan project',
    briefPromptLabel: 'Marketing brief:',
    buildSummary: 'plan, budget ({total} total) and overview deck',
    defaultBudget: MARKETING_DEFAULT_BUDGET,
    sections: [
      {
        heading: 'Audience',
        body:
          '{brief}\n\n' +
          'The audience is defined by who buys, what they need, and where they look first. ' +
          'Campaigns speak to that audience in those places, not to everyone at once.',
      },
      {
        heading: 'Positioning',
        body: 'The offer is positioned on the strength customers already value, stated in one clear line and used consistently across every channel.',
      },
      {
        heading: 'Channels',
        body:
          'Channels are chosen where the audience already is. Each channel has one job — reach, enquiries, or sales — ' +
          'and the budget by channel and month is set out in the budget summary below.',
      },
      {
        heading: 'Campaigns',
        body:
          'Campaigns run in sequence: launch, follow-up, and reminder. ' +
          'Each campaign uses the positioning, points to one clear next step, and is checked against the measures of success.',
      },
      {
        heading: 'Timeline',
        body:
          'Month 1 establishes the channels and the first campaign. ' +
          'Month 2 builds on what worked, shifts budget toward the strongest channel, and drops what did not perform.',
      },
      {
        heading: 'Measures of success',
        body:
          'Success is measured by reach, enquiries, sales, and cost per result, reviewed monthly against the campaign budget. ' +
          'The budget spreadsheet in this project is the source of truth for every figure in this plan and in the overview deck.',
      },
      {
        heading: 'Budget summary',
        body:
          'Campaign budget by channel and month:\n{budgetLines}\n\nTotal campaign budget: {total}. ' +
          'The budget spreadsheet in this project is the source of truth for every figure in this plan.',
      },
    ],
    slides: [
      {
        heading: '{name}',
        bullets: ['A marketing plan, a budget, and an overview — in one project.'],
      },
      {
        heading: 'Audience and positioning',
        bullets: [
          'A defined audience, where they already look.',
          'One clear positioning line, used consistently.',
        ],
      },
      {
        heading: 'Channels and campaigns',
        bullets: [
          'Each channel has one job: reach, enquiries, or sales.',
          'Campaigns run launch, follow-up, reminder.',
        ],
      },
      { heading: 'Campaign budget', bullets: ['Total campaign budget: {total}.', '{budget}'] },
      {
        heading: 'Next steps',
        bullets: [
          'Confirm the channels and the first campaign.',
          'Open the plan document and the budget to review the details.',
        ],
      },
    ],
  },
}

export function projectAiTypeDefinition(projectType: unknown): ProjectAiTypeDefinition {
  return PROJECT_AI_TYPE_DEFINITIONS[normalizeProjectType(projectType)]
}

export function projectTitleFromBrief(brief: string): string {
  const first = brief.trim().split(/[.\n]/)[0] ?? ''
  const cleaned = first.trim()
  if (!cleaned) return 'New Project'
  return cleaned.length > 80 ? `${cleaned.slice(0, 77).trim()}…` : cleaned
}

export function businessNameFromBrief(brief: string): string {
  const first = brief.trim().split(/[.\n]/)[0] ?? ''
  const cleaned = first.trim()
  if (!cleaned) return 'New Business'
  return cleaned.length > 80 ? `${cleaned.slice(0, 77).trim()}…` : cleaned
}

interface TemplateContext {
  name: string
  brief: string
  total: string
  budgetLines: string
}

function fillTemplate(template: string, context: TemplateContext): string {
  return template
    .split('{name}')
    .join(context.name)
    .split('{brief}')
    .join(context.brief)
    .split('{total}')
    .join(context.total)
    .split('{budgetLines}')
    .join(context.budgetLines)
}

/** Deterministic content for any project type. Used whenever the model
 *  is unavailable, and as the shape the model's draft is validated
 *  against. Every type's fallback passes the same bounded validation the
 *  service applies to model drafts. */
export function buildProjectContent(projectType: unknown, brief: string): ProjectDraft {
  const definition = projectAiTypeDefinition(projectType)
  const name =
    definition.id === 'business-plan' ? businessNameFromBrief(brief) : projectTitleFromBrief(brief)
  const budget = definition.defaultBudget.map((row) => ({ ...row }))
  const context: TemplateContext = {
    name,
    brief: brief.trim(),
    total: formatMoney(budgetTotal(budget)),
    budgetLines: budget.map((row) => `- ${row.label}: ${formatMoney(row.amount)}`).join('\n'),
  }
  return {
    projectType: definition.id,
    title: name,
    businessName: name,
    sections: definition.sections.map((section) => ({
      heading: section.heading,
      body: fillTemplate(section.body, context),
    })),
    budget,
  }
}

/** Deterministic business-plan content (the original project type). */
export function buildBusinessPlanContent(brief: string): BusinessPlanDraft {
  return buildProjectContent('business-plan', brief)
}

export function buildGrantApplicationContent(brief: string): ProjectDraft {
  return buildProjectContent('grant-application', brief)
}

export function buildEventPlanContent(brief: string): ProjectDraft {
  return buildProjectContent('event-plan', brief)
}

export function buildMarketingPlanContent(brief: string): ProjectDraft {
  return buildProjectContent('marketing-plan', brief)
}

function isDraftRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

/**
 * The bounded validation every draft must pass, whatever its type or
 * origin: at most 8 sections and 20 budget rows, every section with a
 * heading and a body, every budget row with a label and a positive,
 * finite amount. Returns the normalized draft, or null when invalid.
 * The template fallback for each type passes this, and a model draft
 * that fails it is discarded in favour of that fallback.
 */
export function validateProjectDraft(
  value: unknown,
  brief: string,
  projectType: unknown,
): ProjectDraft | null {
  const type = normalizeProjectType(projectType)
  if (!isDraftRecord(value)) return null
  const sectionsRaw = value.sections
  const budgetRaw = value.budget
  if (!Array.isArray(sectionsRaw) || !Array.isArray(budgetRaw)) return null
  const sections: ProjectDraft['sections'] = []
  for (const entry of sectionsRaw.slice(0, 8)) {
    if (!isDraftRecord(entry)) return null
    const heading = typeof entry.heading === 'string' ? entry.heading.trim() : ''
    const body = typeof entry.body === 'string' ? entry.body.trim() : ''
    if (!heading || !body) return null
    sections.push({ heading, body })
  }
  const budget: BudgetRow[] = []
  for (const entry of budgetRaw.slice(0, 20)) {
    if (!isDraftRecord(entry)) return null
    const label = typeof entry.label === 'string' ? entry.label.trim() : ''
    const amount = typeof entry.amount === 'number' ? entry.amount : Number(entry.amount)
    if (!label || !Number.isFinite(amount) || amount <= 0) return null
    budget.push({ label, amount })
  }
  if (sections.length === 0 || budget.length === 0) return null
  const rawTitle =
    typeof value.title === 'string' && value.title.trim()
      ? value.title.trim()
      : typeof value.businessName === 'string' && value.businessName.trim()
        ? value.businessName.trim()
        : ''
  const title =
    rawTitle ||
    (type === 'business-plan' ? businessNameFromBrief(brief) : projectTitleFromBrief(brief))
  return { projectType: type, title, businessName: title, sections, budget }
}

/** Carry a draft of any origin (template, model, or a pre-types draft
 *  with no type) into the current draft shape for `projectType`. */
export function normalizeProjectDraft(
  draft: Omit<ProjectDraft, 'projectType' | 'title'> & {
    projectType?: unknown
    title?: string
  },
  projectType: unknown,
): ProjectDraft {
  const type = normalizeProjectType(projectType ?? draft.projectType)
  const title = draft.title?.trim() ? draft.title.trim() : draft.businessName
  return { ...draft, projectType: type, title, businessName: title }
}

/** Markdown for the project document, for any project type. */
export function projectDraftMarkdown(draft: ProjectDraft): string {
  const definition = projectAiTypeDefinition(draft.projectType)
  const lines = [`# ${definition.docTitle}: ${draft.businessName}`, '']
  for (const section of draft.sections) {
    lines.push(`## ${section.heading}`, '', section.body, '')
  }
  return lines.join('\n')
}

/** Markdown for the plan document. */
export function businessPlanMarkdown(draft: BusinessPlanDraft): string {
  return projectDraftMarkdown(draft)
}

/** Row matrix for the budget spreadsheet (header, rows, total formula). */
export function projectDraftBudgetRows(draft: ProjectDraft): (string | number)[][] {
  const rows: (string | number)[][] = [['Category', 'Amount']]
  draft.budget.forEach((row) => rows.push([row.label, row.amount]))
  rows.push(['Total', `=SUM(B2:B${draft.budget.length + 1})`])
  return rows
}

/** Row matrix for the budget spreadsheet (header, rows, total formula). */
export function businessPlanBudgetRows(draft: BusinessPlanDraft): (string | number)[][] {
  return projectDraftBudgetRows(draft)
}

/** Markdown outline for the deck (one `#` title per slide), any type. */
export function projectDraftDeckOutline(draft: ProjectDraft): string {
  const definition = projectAiTypeDefinition(draft.projectType)
  const context: TemplateContext = {
    name: draft.businessName,
    brief: '',
    total: formatMoney(budgetTotal(draft.budget)),
    budgetLines: '',
  }
  const lines: string[] = []
  for (const slide of definition.slides) {
    if (lines.length > 0) lines.push('')
    lines.push(`# ${fillTemplate(slide.heading, context)}`)
    for (const bullet of slide.bullets) {
      if (bullet === '{budget}') {
        for (const row of draft.budget) lines.push(`- ${row.label}: ${formatMoney(row.amount)}.`)
      } else {
        lines.push(`- ${fillTemplate(bullet, context)}`)
      }
    }
  }
  return lines.join('\n')
}

/** Markdown outline for the pitch deck (one `#` title per slide). */
export function businessPlanDeckOutline(draft: BusinessPlanDraft): string {
  return projectDraftDeckOutline(draft)
}
