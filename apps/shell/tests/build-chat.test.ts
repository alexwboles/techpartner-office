import { describe, expect, it } from 'vitest'
import {
  applyBuildFill,
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
  type BuildPlanItem,
} from '../src/shared/build-chat'
import { TEMPLATES, type TemplateDef } from '../src/shared/template-defs'

const invoice = TEMPLATES.find((template) => template.id === 'invoice-service')!
const resume = TEMPLATES.find((template) => template.id === 'resume-modern')!

const customTemplates: TemplateDef[] = [
  {
    id: 'letter',
    title: 'Formal Letter',
    description: 'A formal business letter with sender, date and signature.',
    category: 'business-letters',
    content: {
      app: 'docs',
      html: '<h1>[Recipient Name]</h1><p>Dear [Recipient Name],</p><p>About [Subject].</p>',
    },
  },
  {
    id: 'grid',
    title: 'Simple Grid',
    description: 'A plain grid.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: '[Client] Budget',
          headerRow: true,
          rows: [
            ['Item', 'Cost'],
            ['[Item name]', { formula: 'B2', cached: '[Item name] total' }],
          ],
        },
      ],
    },
  },
  {
    id: 'deck',
    title: 'Simple Deck',
    description: 'A plain deck.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        { title: '[Talk Title]', subtitle: 'By [Speaker]', bullets: ['For [Audience]'] },
        { title: 'Agenda', bullets: ['One'] },
      ],
    },
  },
]

describe('parseBuildReply: questions shape', () => {
  it('accepts a questions reply and keeps the message', () => {
    const reply = parseBuildReply(
      {
        questions: ['Who is the invoice for?', 'What are the line items?'],
        message: 'Two things first.',
      },
      TEMPLATES,
    )
    expect(reply).toEqual({
      kind: 'questions',
      message: 'Two things first.',
      questions: ['Who is the invoice for?', 'What are the line items?'],
    })
  })

  it('caps the question count and drops blanks', () => {
    const reply = parseBuildReply({ questions: [' a ', '', 'b', 'c', 'd', 'e', 'f'] }, TEMPLATES)
    expect(reply?.kind).toBe('questions')
    if (reply?.kind === 'questions') {
      expect(reply.questions).toEqual(['a', 'b', 'c', 'd', 'e'])
    }
  })

  it('rejects an empty questions array', () => {
    expect(parseBuildReply({ questions: [] }, TEMPLATES)).toBeNull()
    expect(parseBuildReply({ questions: ['', '  '] }, TEMPLATES)).toBeNull()
  })

  it('a questions reply can never carry plan items', () => {
    const reply = parseBuildReply(
      {
        questions: ['Which client?'],
        plan: [{ kind: 'create', app: 'docs', title: 'X', markdown: '# X' }],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('questions')
  })
})

describe('parseBuildReply: plan shape', () => {
  it('accepts a template create with fill', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'create',
            app: 'sheets',
            title: 'Invoice for Acme',
            templateId: 'invoice-service',
            fill: [{ find: '[Your Business Name]', replace: 'Acme' }],
          },
        ],
        message: 'Making your invoice.',
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
    if (reply?.kind === 'plan') {
      expect(reply.items).toHaveLength(1)
      expect(reply.items[0]).toMatchObject({
        kind: 'create',
        app: 'sheets',
        templateId: 'invoice-service',
      })
    }
  })

  it('accepts a multi-app custom plan in one reply', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'create',
            app: 'docs',
            title: 'Cover letter',
            markdown: '# Cover letter\n\nDear…',
          },
          {
            kind: 'create',
            app: 'sheets',
            title: 'Quote',
            sheets: [
              {
                name: 'Quote',
                headerRow: true,
                rows: [
                  ['Item', 'Price'],
                  ['Work', 100],
                ],
              },
            ],
          },
          {
            kind: 'create',
            app: 'slides',
            title: 'Pitch',
            slides: [{ title: 'Pitch', bullets: ['One'] }],
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
    if (reply?.kind === 'plan') {
      expect(reply.items.map((item) => (item.kind === 'create' ? item.app : 'edit'))).toEqual([
        'docs',
        'sheets',
        'slides',
      ])
    }
  })

  it('accepts formula cells in custom sheets', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'create',
            app: 'sheets',
            title: 'Calc',
            sheets: [
              {
                name: 'Calc',
                rows: [
                  ['Qty', 'Rate', 'Amount'],
                  [2, 50, { formula: 'A2*B2', cached: 100, format: 'currency' }],
                ],
              },
            ],
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
  })

  it('accepts an edit item for a document', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'edit',
            app: 'docs',
            path: '/Users/alex/Documents/letter.docx',
            find: 'Dear Sir',
            replace: 'Dear Ms. Lee',
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
    if (reply?.kind === 'plan') expect(planHasEdits(reply.items)).toBe(true)
  })

  it('accepts a Windows path on an edit item', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'edit',
            app: 'slides',
            path: 'C:\\Users\\alex\\deck.pptx',
            find: 'Q3',
            replace: 'Q4',
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
  })
})

describe('parseBuildReply: sheet-edit items', () => {
  it('accepts a sheet edit and normalizes the cell references', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'sheet-edit',
            path: 'Budget.xlsx',
            updates: [
              { cell: 'b12', value: 4500 },
              { cell: 'C3', value: 'Total' },
            ],
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
    if (reply?.kind === 'plan') {
      expect(planHasEdits(reply.items)).toBe(true)
      expect(reply.items[0]).toEqual({
        kind: 'sheet-edit',
        path: 'Budget.xlsx',
        updates: [
          { cell: 'B12', value: 4500 },
          { cell: 'C3', value: 'Total' },
        ],
      })
    }
  })

  it('accepts a named sheet and a subfolder path', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'sheet-edit',
            path: 'Finance/Q3.xlsx',
            sheet: 'Totals',
            updates: [{ cell: 'A1', value: 1 }],
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
  })

  const invalidSheetEdits: Array<[string, unknown]> = [
    ['an absolute path', { path: '/home/alex/Budget.xlsx', updates: [{ cell: 'A1', value: 1 }] }],
    [
      'a Windows absolute path',
      { path: 'C:\\Users\\alex\\Budget.xlsx', updates: [{ cell: 'A1', value: 1 }] },
    ],
    [
      'a path climbing out of the folder',
      { path: '../Budget.xlsx', updates: [{ cell: 'A1', value: 1 }] },
    ],
    ['a non-spreadsheet file', { path: 'Notes.docx', updates: [{ cell: 'A1', value: 1 }] }],
    ['a reversed ref', { path: 'Budget.xlsx', updates: [{ cell: '12B', value: 1 }] }],
    ['row zero', { path: 'Budget.xlsx', updates: [{ cell: 'B0', value: 1 }] }],
    ['no row at all', { path: 'Budget.xlsx', updates: [{ cell: 'B', value: 1 }] }],
    ['a column past XFD', { path: 'Budget.xlsx', updates: [{ cell: 'XFE1', value: 1 }] }],
    [
      'a formula-looking value',
      { path: 'Budget.xlsx', updates: [{ cell: 'B2', value: '=SUM(A1:A9)' }] },
    ],
    [
      'the same cell twice',
      {
        path: 'Budget.xlsx',
        updates: [
          { cell: 'B2', value: 1 },
          { cell: 'b2', value: 2 },
        ],
      },
    ],
    ['no updates', { path: 'Budget.xlsx', updates: [] }],
    [
      'more than 50 updates',
      {
        path: 'Budget.xlsx',
        updates: Array.from({ length: 51 }, (_, i) => ({ cell: `A${i + 1}`, value: i })),
      },
    ],
    ['a boolean value', { path: 'Budget.xlsx', updates: [{ cell: 'B2', value: true }] }],
  ]
  it.each(invalidSheetEdits)('rejects %s', (_label, item) => {
    const reply = { plan: [{ kind: 'sheet-edit', ...(item as Record<string, unknown>) }] }
    expect(parseBuildReply(reply, TEMPLATES)).toBeNull()
  })
})

describe('parseBuildReply: derive items', () => {
  it('accepts a derive item and keeps it out of the edit gate', () => {
    const reply = parseBuildReply(
      {
        plan: [
          {
            kind: 'derive',
            source: 'Budget.xlsx',
            instruction: 'Summarize it as a one-page brief',
            title: 'Budget Summary',
          },
        ],
      },
      TEMPLATES,
    )
    expect(reply?.kind).toBe('plan')
    if (reply?.kind === 'plan') {
      // a derive creates a new file: no confirmation gate
      expect(planHasEdits(reply.items)).toBe(false)
      expect(reply.items[0]).toEqual({
        kind: 'derive',
        source: 'Budget.xlsx',
        instruction: 'Summarize it as a one-page brief',
        title: 'Budget Summary',
      })
    }
  })

  const invalidDerives: Array<[string, unknown]> = [
    [
      'an absolute source',
      { source: '/home/alex/Notes.docx', instruction: 'Sum it up', title: 'S' },
    ],
    ['a source climbing out', { source: '../Notes.docx', instruction: 'Sum it up', title: 'S' }],
    ['an unsupported type', { source: 'Scan.pdf', instruction: 'Sum it up', title: 'S' }],
    ['no instruction', { source: 'Notes.docx', instruction: '', title: 'S' }],
    ['no title', { source: 'Notes.docx', instruction: 'Sum it up', title: '' }],
  ]
  it.each(invalidDerives)('rejects %s', (_label, item) => {
    const reply = { plan: [{ kind: 'derive', ...(item as Record<string, unknown>) }] }
    expect(parseBuildReply(reply, TEMPLATES)).toBeNull()
  })
})

describe('parseBuildReply: invalid output parses to null', () => {
  const invalid: Array<[string, unknown]> = [
    ['plain prose', 'Sure, making that now!'],
    ['an array', [{ questions: ['x'] }]],
    ['an empty object', {}],
    ['an empty plan', { plan: [] }],
    [
      'too many items',
      {
        plan: Array.from({ length: 7 }, () => ({
          kind: 'create',
          app: 'docs',
          title: 'T',
          markdown: '# T',
        })),
      },
    ],
    [
      'unknown template',
      { plan: [{ kind: 'create', app: 'docs', title: 'T', templateId: 'nope' }] },
    ],
    [
      'template for the wrong app',
      { plan: [{ kind: 'create', app: 'docs', title: 'T', templateId: 'invoice-service' }] },
    ],
    ['docs without content', { plan: [{ kind: 'create', app: 'docs', title: 'T' }] }],
    [
      'sheets without rows',
      { plan: [{ kind: 'create', app: 'sheets', title: 'T', sheets: [{ name: 'S', rows: [] }] }] },
    ],
    [
      'a bad sheet name',
      {
        plan: [
          { kind: 'create', app: 'sheets', title: 'T', sheets: [{ name: 'a/b', rows: [['x']] }] },
        ],
      },
    ],
    [
      'slides without a title',
      { plan: [{ kind: 'create', app: 'slides', title: 'T', slides: [{ bullets: [] }] }] },
    ],
    ['a pdf item', { plan: [{ kind: 'create', app: 'pdf', title: 'T', markdown: '# T' }] }],
    [
      'an edit with a relative path',
      { plan: [{ kind: 'edit', app: 'docs', path: 'docs/letter.docx', find: 'a', replace: 'b' }] },
    ],
    [
      'an edit on a spreadsheet',
      { plan: [{ kind: 'edit', app: 'sheets', path: '/tmp/a.xlsx', find: 'a', replace: 'b' }] },
    ],
    [
      'an edit that changes nothing',
      { plan: [{ kind: 'edit', app: 'docs', path: '/tmp/a.docx', find: 'same', replace: 'same' }] },
    ],
    [
      'one bad item poisons the plan',
      {
        plan: [
          { kind: 'create', app: 'docs', title: 'Good', markdown: '# Good' },
          { kind: 'create', app: 'docs', title: 'Bad' },
        ],
      },
    ],
  ]
  it.each(invalid)('%s', (_label, value) => {
    expect(parseBuildReply(value, TEMPLATES)).toBeNull()
  })
})

describe('describePlan', () => {
  it('describes template, custom and edit items in plain language', () => {
    const items: BuildPlanItem[] = [
      { kind: 'create', app: 'sheets', title: 'Invoice for Acme', templateId: 'invoice-service' },
      { kind: 'create', app: 'docs', title: 'Notes', markdown: '# Notes' },
      {
        kind: 'edit',
        app: 'docs',
        path: '/home/alex/letter.docx',
        find: 'Dear Sir',
        replace: 'Dear Ms. Lee',
      },
      {
        kind: 'sheet-edit',
        path: 'Budget.xlsx',
        updates: [
          { cell: 'B12', value: 4500 },
          { cell: 'C3', value: 'Total' },
        ],
      },
      {
        kind: 'derive',
        source: 'Finance/Budget.xlsx',
        instruction: 'Summarize it',
        title: 'Budget Summary',
      },
    ]
    expect(describePlan(items, TEMPLATES)).toEqual([
      'Create a spreadsheet “Invoice for Acme” from the “Service Invoice” template',
      'Create a document “Notes”',
      'Edit letter.docx: replace “Dear Sir” with “Dear Ms. Lee”',
      'Edit Budget.xlsx: set B12 to 4500, set C3 to “Total”',
      'Create a document “Budget Summary” from Budget.xlsx',
    ])
  })
})

describe('isBuildConfirmation', () => {
  it.each(['yes', 'Yes.', 'ok', 'go ahead', 'do it', 'yes please'])('accepts %j', (text) => {
    expect(isBuildConfirmation(text)).toBe(true)
  })
  it.each(['no', 'yes but change the title first and also the date', 'maybe', ''])(
    'rejects %j',
    (text) => {
      expect(isBuildConfirmation(text)).toBe(false)
    },
  )
})

describe('template placeholders + fill', () => {
  it('extracts placeholders from every app surface, de-duplicated', () => {
    expect(extractTemplatePlaceholders(customTemplates[0]!)).toEqual([
      '[Recipient Name]',
      '[Subject]',
    ])
    expect(extractTemplatePlaceholders(customTemplates[1]!)).toEqual(['[Client]', '[Item name]'])
    expect(extractTemplatePlaceholders(customTemplates[2]!)).toEqual([
      '[Talk Title]',
      '[Speaker]',
      '[Audience]',
    ])
  })

  it('extracts real placeholders from the shipped gallery', () => {
    const placeholders = extractTemplatePlaceholders(resume)
    expect(placeholders).toContain('[Your Full Name]')
    expect(placeholders.length).toBeGreaterThan(3)
  })

  it('applies fill across docs HTML without mutating the template', () => {
    const filled = applyBuildFill(customTemplates[0]!, [
      { find: '[Recipient Name]', replace: 'Ms. Lee' },
      { find: '[Subject]', replace: 'the quote' },
    ])
    expect(filled.content.app).toBe('docs')
    if (filled.content.app === 'docs') {
      expect(filled.content.html).toContain('<h1>Ms. Lee</h1>')
      expect(filled.content.html).toContain('Dear Ms. Lee,')
      expect(filled.content.html).toContain('About the quote.')
    }
    const original = customTemplates[0]!
    if (original.content.app === 'docs') {
      expect(original.content.html).toContain('[Recipient Name]')
    }
  })

  it('applies fill to sheet names, string cells and cached string results', () => {
    const filled = applyBuildFill(customTemplates[1]!, [
      { find: '[Client]', replace: 'Acme' },
      { find: '[Item name]', replace: 'Design work' },
    ])
    if (filled.content.app === 'sheets') {
      const sheet = filled.content.sheets[0]!
      expect(sheet.name).toBe('Acme Budget')
      expect(sheet.rows[1]![0]).toBe('Design work')
      const formula = sheet.rows[1]![1]
      expect(typeof formula).toBe('object')
      if (typeof formula === 'object') {
        expect(formula.formula).toBe('B2')
        expect(formula.cached).toBe('Design work total')
      }
    }
  })

  it('applies fill to slide titles, subtitles and bullets', () => {
    const filled = applyBuildFill(customTemplates[2]!, [
      { find: '[Talk Title]', replace: 'Q4 Review' },
      { find: '[Speaker]', replace: 'Alex' },
      { find: '[Audience]', replace: 'the team' },
    ])
    if (filled.content.app === 'slides') {
      expect(filled.content.slides[0]).toEqual({
        title: 'Q4 Review',
        subtitle: 'By Alex',
        bullets: ['For the team'],
      })
    }
  })
})

describe('deterministic fallback', () => {
  it('matches a clear template request', () => {
    expect(matchTemplateForMessage('I need a service invoice', TEMPLATES)?.id).toBe(
      'invoice-service',
    )
    expect(matchTemplateForMessage('make me a modern resume', TEMPLATES)?.id).toBe('resume-modern')
  })

  it('returns null when nothing clearly matches', () => {
    expect(matchTemplateForMessage('blorp zanzibar quixotic', TEMPLATES)).toBeNull()
    expect(matchTemplateForMessage('', TEMPLATES)).toBeNull()
  })

  it('breaks a tie toward the one template that asks questions first', () => {
    // "invoice" scores both invoice templates equally; the product one
    // carries [Product n] placeholders while the service one would build
    // instantly with sample line items — clarifying beats guessing
    expect(matchTemplateForMessage('Make an invoice', TEMPLATES)?.id).toBe('invoice-product-sales')
    // a tie where several templates ask questions is still a guess
    expect(matchTemplateForMessage('Create a presentation', TEMPLATES)).toBeNull()
  })

  it('asks one question per placeholder', () => {
    expect(fallbackQuestionsFor(['[Name]', '[Date]'])).toEqual([
      'What should I put for [Name]?',
      'What should I put for [Date]?',
    ])
  })

  it('parses a single-placeholder answer as the whole message', () => {
    const answers = parsePlaceholderAnswers(['[Your Full Name]'], 'Alex Boles')
    expect(answers?.get('[Your Full Name]')).toBe('Alex Boles')
  })

  it('parses labelled multi-placeholder answers', () => {
    const answers = parsePlaceholderAnswers(
      ['[Name]', '[Date]'],
      '[Name]: Alex Boles\n[Date]: Friday',
    )
    expect(answers?.get('[Name]')).toBe('Alex Boles')
    expect(answers?.get('[Date]')).toBe('Friday')
  })

  it('parses positional multi-placeholder answers, one per line', () => {
    const answers = parsePlaceholderAnswers(['[Name]', '[Date]'], 'Alex Boles\nFriday')
    expect(answers?.get('[Name]')).toBe('Alex Boles')
    expect(answers?.get('[Date]')).toBe('Friday')
  })

  it('refuses an incomplete answer set', () => {
    expect(parsePlaceholderAnswers(['[Name]', '[Date]'], 'Alex Boles')).toBeNull()
    expect(parsePlaceholderAnswers(['[Name]'], '   ')).toBeNull()
  })

  it('builds the fallback create item with its fill', () => {
    const answers = parsePlaceholderAnswers(['[Your Full Name]'], 'Alex Boles')
    const item = fallbackCreateItem(resume, answers)
    expect(item).toMatchObject({
      kind: 'create',
      app: 'docs',
      templateId: 'resume-modern',
      fill: [{ find: '[Your Full Name]', replace: 'Alex Boles' }],
    })
  })
})

describe('buildChatSystemPrompt', () => {
  it('carries the contract: JSON shapes, clarify first, no PDF, the catalog', () => {
    const prompt = buildChatSystemPrompt(TEMPLATES)
    expect(prompt).toContain('"questions"')
    expect(prompt).toContain('"plan"')
    expect(prompt).toContain('cannot create PDF')
    expect(prompt).toContain('invoice-service (sheets)')
    expect(prompt).toContain(invoice.title)
  })

  it('teaches the spreadsheet edit and its literal-values rule', () => {
    const prompt = buildChatSystemPrompt(TEMPLATES)
    expect(prompt).toContain('"sheet-edit"')
    expect(prompt).toContain('never start a value with "="')
    expect(prompt).not.toContain('Spreadsheets cannot be edited')
  })

  it('teaches derive, and the writing prompt forbids inventing facts', () => {
    const prompt = buildChatSystemPrompt(TEMPLATES)
    expect(prompt).toContain('"derive"')
    expect(prompt).toContain('never changed')
    const writing = deriveDocumentSystemPrompt()
    expect(writing).toContain('markdown only')
    expect(writing).toContain('Do not invent')
  })
})
