import { describe, expect, it } from 'vitest'
import {
  PROJECT_AI_PROJECT_TYPES,
  buildBusinessPlanContent,
  buildEventPlanContent,
  buildGrantApplicationContent,
  buildMarketingPlanContent,
  buildProjectContent,
  businessNameFromBrief,
  businessPlanBudgetRows,
  businessPlanDeckOutline,
  businessPlanMarkdown,
  budgetTotal,
  findNumberSyncs,
  findTypos,
  formatMoney,
  isSafeSuggestionClass,
  mayAutoApply,
  normalizeProjectAiState,
  normalizeProjectType,
  projectAiFileName,
  projectDraftBudgetRows,
  projectDraftDeckOutline,
  projectDraftMarkdown,
  projectTypeOfState,
  reviewProjectContent,
  validateModelSuggestions,
  validateProjectDraft,
  type ProjectAiProjectType,
  type ProjectAiSuggestion,
} from '../src/shared/project-ai'

const BUDGET = [
  { label: 'Equipment', amount: 8500 },
  { label: 'Marketing launch', amount: 3500 },
  { label: 'Working capital reserve', amount: 10000 },
]

describe('SAFE vs queue classification', () => {
  it('only typo and number-sync classes are SAFE classes', () => {
    expect(isSafeSuggestionClass('typo')).toBe(true)
    expect(isSafeSuggestionClass('number-sync')).toBe(true)
    expect(isSafeSuggestionClass('other')).toBe(false)
  })

  it('a language-model suggestion is never auto-appliable', () => {
    const modelSuggestion: ProjectAiSuggestion = {
      id: 'other-doc-1',
      kind: 'other',
      file: 'doc',
      fileName: 'business-plan.docx',
      title: 'Tighten the executive summary',
      detail: 'Suggested by the model.',
      before: 'This plan sets out',
      after: 'This plan shows',
      autoSafe: false,
    }
    expect(mayAutoApply(modelSuggestion)).toBe(false)
    // even a mis-flagged model suggestion must not pass the gate
    expect(mayAutoApply({ ...modelSuggestion, autoSafe: true })).toBe(false)
  })

  it('a SAFE suggestion needs a non-empty exact replacement target', () => {
    const empty: ProjectAiSuggestion = {
      id: 'typo-doc-1',
      kind: 'typo',
      file: 'doc',
      fileName: 'business-plan.docx',
      title: 'Typo',
      detail: '',
      before: '',
      after: 'the',
      autoSafe: true,
    }
    expect(mayAutoApply(empty)).toBe(false)
  })
})

describe('typo findings', () => {
  it('finds whole-word typos and preserves capitalisation', () => {
    const found = findTypos('Teh buisness plan is ready.', 'doc', 0)
    expect(found).toHaveLength(2)
    expect(found[0]).toMatchObject({ kind: 'typo', before: 'Teh', after: 'The', autoSafe: true })
    expect(found[1]).toMatchObject({ before: 'buisness', after: 'business', autoSafe: true })
  })

  it('does not flag substrings inside correct words', () => {
    expect(findTypos('The theatre business is open.', 'doc', 0)).toHaveLength(0)
  })
})

describe('number syncs', () => {
  it('corrects a doc figure to the spreadsheet value when exactly one row matches', () => {
    const found = findNumberSyncs('Equipment: $9,000 for the launch.', 'doc', BUDGET, 0)
    expect(found).toHaveLength(1)
    expect(found[0]).toMatchObject({
      kind: 'number-sync',
      before: '$9,000',
      after: '$8,500',
      autoSafe: true,
    })
  })

  it('leaves a matching figure alone', () => {
    expect(findNumberSyncs('Equipment: $8,500 for the launch.', 'doc', BUDGET, 0)).toHaveLength(0)
  })

  it('does not correct a figure that already matches one of several matching rows', () => {
    const rows = [
      { label: 'Marketing', amount: 3500 },
      { label: 'Marketing launch', amount: 4000 },
    ]
    // $3,500 already matches the Marketing row, so there is nothing to
    // fix — correcting it to $4,000 would invent a discrepancy.
    expect(findNumberSyncs('Marketing launch budget: $3,500.', 'deck', rows, 0)).toHaveLength(0)
  })

  it('queues (never auto-applies) a figure matching several budget rows', () => {
    const rows = [
      { label: 'Marketing', amount: 3500 },
      { label: 'Marketing launch', amount: 4000 },
    ]
    const found = findNumberSyncs('Marketing launch budget: $9,000.', 'deck', rows, 0)
    expect(found).toHaveLength(1)
    expect(found[0].autoSafe).toBe(false)
    expect(found[0].ambiguity).toMatch(/2 budget rows/)
    expect(mayAutoApply(found[0])).toBe(false)
  })
})

describe('cross-file review', () => {
  it('splits SAFE findings from the approval queue', () => {
    const { safe, queued } = reviewProjectContent({
      docText: 'Teh plan needs Equipment: $9,000.',
      deckText: 'Total startup costs: $50,000.',
      budgetRows: BUDGET,
    })
    expect(safe.map((s) => s.kind).sort()).toEqual(['number-sync', 'typo'])
    expect(queued).toHaveLength(0)
  })

  it('formats money without cents for whole amounts', () => {
    expect(formatMoney(8500)).toBe('$8,500')
    expect(formatMoney(3500.5)).toBe('$3,500.50')
  })
})

describe('business-plan content', () => {
  it('derives a business name and a consistent budget', () => {
    expect(businessNameFromBrief('A bakery in Palatka. Fresh bread daily.')).toBe(
      'A bakery in Palatka',
    )
    const draft = buildBusinessPlanContent('A bakery in Palatka. Fresh bread daily.')
    expect(draft.businessName).toBe('A bakery in Palatka')
    expect(budgetTotal(draft.budget)).toBeGreaterThan(0)
    const markdown = businessPlanMarkdown(draft)
    expect(markdown).toContain('Financial summary')
    expect(markdown).toContain(formatMoney(budgetTotal(draft.budget)))
    const rows = businessPlanBudgetRows(draft)
    expect(rows[0]).toEqual(['Category', 'Amount'])
    expect(rows.at(-1)?.[0]).toBe('Total')
    const outline = businessPlanDeckOutline(draft)
    expect(outline).toContain('# Financials')
    expect(outline).toContain(formatMoney(budgetTotal(draft.budget)))
  })
})

describe('project types', () => {
  it('lists the business plan plus the three new types, business plan first', () => {
    expect(PROJECT_AI_PROJECT_TYPES).toEqual([
      'business-plan',
      'grant-application',
      'event-plan',
      'marketing-plan',
    ])
    expect(projectAiFileName('doc')).toBe('business-plan.docx')
    expect(projectAiFileName('sheet')).toBe('business-plan-budget.xlsx')
    expect(projectAiFileName('deck')).toBe('business-plan-pitch.pptx')
  })

  it('old state without a type defaults to business-plan', () => {
    expect(normalizeProjectType(undefined)).toBe('business-plan')
    expect(normalizeProjectType('not-a-type')).toBe('business-plan')
    expect(projectTypeOfState({})).toBe('business-plan')
    expect(projectTypeOfState({ projectType: 'event-plan' })).toBe('event-plan')
    const restored = normalizeProjectAiState({ folder: '/tmp/project', status: 'ready' })
    expect(restored.projectType).toBe('business-plan')
    expect(restored.folder).toBe('/tmp/project')
    expect(normalizeProjectAiState({ projectType: 'grant-application' }).projectType).toBe(
      'grant-application',
    )
  })
})

interface NewTypeCase {
  type: Exclude<ProjectAiProjectType, 'business-plan'>
  brief: string
  docTitle: string
  headings: string[]
  budgetLabels: string[]
  docFile: string
  sheetFile: string
  deckFile: string
  deckBudgetHeading: string
  build: (brief: string) => ReturnType<typeof buildProjectContent>
}

const NEW_TYPE_CASES: NewTypeCase[] = [
  {
    type: 'grant-application',
    brief: 'Fund a community garden programme in Palatka. Fresh produce for local families.',
    docTitle: 'Grant application',
    headings: [
      'Need statement',
      'Project description',
      'Goals and outcomes',
      'Evaluation',
      'Organizational capacity',
    ],
    budgetLabels: ['Personnel', 'Supplies', 'Equipment', 'Other costs'],
    docFile: 'grant-application.docx',
    sheetFile: 'grant-application-budget.xlsx',
    deckFile: 'grant-application-overview.pptx',
    deckBudgetHeading: 'Requested funding',
    build: buildGrantApplicationContent,
  },
  {
    type: 'event-plan',
    brief: 'A spring fundraiser dinner at the ranch. Sponsors, volunteers and local guests.',
    docTitle: 'Event plan',
    headings: ['Overview', 'Schedule / run of show', 'Venue & logistics', 'Roles', 'Checklist'],
    budgetLabels: ['Venue', 'Catering', 'Materials', 'Promotion'],
    docFile: 'event-plan.docx',
    sheetFile: 'event-plan-budget.xlsx',
    deckFile: 'event-plan-overview.pptx',
    deckBudgetHeading: 'Event budget',
    build: buildEventPlanContent,
  },
  {
    type: 'marketing-plan',
    brief: 'Launch marketing for a local bakery. Social, search and email over two months.',
    docTitle: 'Marketing plan',
    headings: [
      'Audience',
      'Positioning',
      'Channels',
      'Campaigns',
      'Timeline',
      'Measures of success',
    ],
    budgetLabels: [
      'Social media - Month 1',
      'Search advertising - Month 1',
      'Content creation - Month 1',
      'Social media - Month 2',
      'Search advertising - Month 2',
      'Email marketing - Month 2',
    ],
    docFile: 'marketing-plan.docx',
    sheetFile: 'marketing-plan-budget.xlsx',
    deckFile: 'marketing-plan-overview.pptx',
    deckBudgetHeading: 'Campaign budget',
    build: buildMarketingPlanContent,
  },
]

for (const testCase of NEW_TYPE_CASES) {
  describe(`project type: ${testCase.type}`, () => {
    const draft = testCase.build(testCase.brief)

    it('fallback draft carries the type and passes draft validation', () => {
      expect(draft.projectType).toBe(testCase.type)
      expect(buildProjectContent(testCase.type, testCase.brief)).toEqual(draft)
      expect(validateProjectDraft(draft, testCase.brief, testCase.type)).toEqual(draft)
      expect(draft.sections.length).toBeGreaterThan(0)
      expect(draft.sections.length).toBeLessThanOrEqual(8)
      for (const section of draft.sections) {
        expect(section.heading).not.toBe('')
        expect(section.body).not.toBe('')
      }
      expect(draft.budget.length).toBeLessThanOrEqual(20)
      for (const row of draft.budget) {
        expect(Number.isFinite(row.amount)).toBe(true)
        expect(row.amount).toBeGreaterThan(0)
      }
      // a draft with a non-positive amount fails the same validation
      expect(
        validateProjectDraft(
          { ...draft, budget: [{ label: 'Free', amount: 0 }] },
          testCase.brief,
          testCase.type,
        ),
      ).toBeNull()
    })

    it('has the required sections and budget shape', () => {
      const headings = draft.sections.map((section) => section.heading)
      for (const heading of testCase.headings) expect(headings).toContain(heading)
      expect(draft.budget.map((row) => row.label)).toEqual(testCase.budgetLabels)
      expect(budgetTotal(draft.budget)).toBeGreaterThan(0)
    })

    it('generates doc, sheet and deck file definitions', () => {
      expect(projectAiFileName('doc', testCase.type)).toBe(testCase.docFile)
      expect(projectAiFileName('sheet', testCase.type)).toBe(testCase.sheetFile)
      expect(projectAiFileName('deck', testCase.type)).toBe(testCase.deckFile)
      const markdown = projectDraftMarkdown(draft)
      expect(markdown).toContain(`# ${testCase.docTitle}:`)
      for (const heading of testCase.headings) expect(markdown).toContain(`## ${heading}`)
      expect(markdown).toContain(formatMoney(budgetTotal(draft.budget)))
      const rows = projectDraftBudgetRows(draft)
      expect(rows[0]).toEqual(['Category', 'Amount'])
      expect(rows).toHaveLength(draft.budget.length + 2)
      expect(rows.at(-1)?.[0]).toBe('Total')
      const outline = projectDraftDeckOutline(draft)
      expect(outline).toContain(`# ${testCase.deckBudgetHeading}`)
      expect(outline).toContain(formatMoney(budgetTotal(draft.budget)))
      for (const row of draft.budget) {
        expect(outline).toContain(`${row.label}: ${formatMoney(row.amount)}`)
      }
    })

    it('syncs an unambiguous doc figure to the budget as SAFE', () => {
      const row = draft.budget[0]
      const found = findNumberSyncs(
        `${row.label}: $99,999 for the work.`,
        'doc',
        draft.budget,
        0,
        testCase.type,
      )
      expect(found).toHaveLength(1)
      expect(found[0]).toMatchObject({
        kind: 'number-sync',
        before: '$99,999',
        after: formatMoney(row.amount),
        autoSafe: true,
        fileName: testCase.docFile,
      })
      expect(mayAutoApply(found[0])).toBe(true)
    })

    it('queues an ambiguous doc figure instead of auto-applying it', () => {
      const first = draft.budget[0]
      const rows = [first, { label: `${first.label} total`, amount: first.amount + 500 }]
      const { safe, queued } = reviewProjectContent({
        docText: `${first.label} total: $99,999.`,
        deckText: '',
        budgetRows: rows,
        projectType: testCase.type,
      })
      expect(safe).toHaveLength(0)
      expect(queued).toHaveLength(1)
      expect(queued[0]).toMatchObject({
        kind: 'number-sync',
        autoSafe: false,
        fileName: testCase.docFile,
      })
      expect(queued[0].ambiguity).toMatch(/2 budget rows/)
      expect(mayAutoApply(queued[0])).toBe(false)
    })

    it('never marks a model suggestion for the type as autoSafe', () => {
      const suggestions = validateModelSuggestions(
        [
          {
            title: 'Fix teh heading',
            detail: 'Looks like a typo, but it came from the model.',
            file: 'doc',
            before: 'teh',
            after: 'the',
          },
          {
            title: 'Correct the figure',
            detail: 'Looks like a number sync, but it came from the model.',
            file: 'deck',
            before: '$99,999',
            after: '$1,000',
          },
        ],
        0,
        testCase.type,
      )
      expect(suggestions).toHaveLength(2)
      for (const suggestion of suggestions) {
        expect(suggestion.kind).toBe('other')
        expect(suggestion.autoSafe).toBe(false)
        expect(mayAutoApply(suggestion)).toBe(false)
        expect(suggestion.fileName).toBe(
          suggestion.file === 'doc' ? testCase.docFile : testCase.deckFile,
        )
      }
    })
  })
}
