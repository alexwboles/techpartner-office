import { describe, expect, it } from 'vitest'
import {
  buildBusinessPlanContent,
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
  reviewProjectContent,
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
