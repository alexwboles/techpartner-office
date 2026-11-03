import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  TEMPLATES,
  TEMPLATE_CATEGORIES,
  templateCategoryKey,
  templateExtOf,
  templatePreview,
  templateVisualPreview,
  type TemplateDef,
} from '../src/shared/template-defs'
import { buildTemplatePptx, buildTemplateXlsx } from '../src/main/template-files'

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

// Mirror the builder's XML escaping: slide titles land in slide XML escaped,
// so a title like 'Team & The Ask' appears as 'Team &amp; The Ask'.
function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

describe('template registry completeness', () => {
  it('has unique ids and covers every category', () => {
    const ids = TEMPLATES.map((template) => template.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect(TEMPLATES.length).toBeGreaterThanOrEqual(40)
    const covered = new Set(TEMPLATES.map((template) => template.category))
    for (const category of TEMPLATE_CATEGORIES) {
      expect(covered.has(category)).toBe(true)
    }
  })

  it('gives every template a name, description, category, app, and content', () => {
    for (const template of TEMPLATES) {
      expect(template.title.trim().length).toBeGreaterThan(0)
      expect(template.description.trim().length).toBeGreaterThan(20)
      expect(TEMPLATE_CATEGORIES).toContain(template.category)
      expect(templateExtOf(template)).toMatch(/^(docx|xlsx|pptx)$/)
    }
  })

  it('gives every docs template real starter content, not an empty page', () => {
    const docs = TEMPLATES.filter(
      (template): template is TemplateDef & { content: { app: 'docs'; html: string } } =>
        template.content.app === 'docs',
    )
    expect(docs.length).toBeGreaterThan(0)
    for (const template of docs) {
      const text = stripHtml(template.content.html)
      expect(text.length).toBeGreaterThan(200)
      expect(template.content.html).toMatch(/<h[12][^>]*>/)
    }
  })

  it('gives every sheets template at least one sheet with rows and a header', () => {
    const sheets = TEMPLATES.filter((template) => template.content.app === 'sheets')
    expect(sheets.length).toBeGreaterThan(0)
    for (const template of sheets) {
      if (template.content.app !== 'sheets') continue
      expect(template.content.sheets.length).toBeGreaterThanOrEqual(1)
      for (const sheet of template.content.sheets) {
        expect(sheet.rows.length).toBeGreaterThanOrEqual(2)
        expect(sheet.rows[0]!.length).toBeGreaterThanOrEqual(2)
      }
    }
  })

  it('gives every slides template at least three slides, each titled', () => {
    const slides = TEMPLATES.filter((template) => template.content.app === 'slides')
    expect(slides.length).toBeGreaterThan(0)
    for (const template of slides) {
      if (template.content.app !== 'slides') continue
      expect(template.content.slides.length).toBeGreaterThanOrEqual(3)
      for (const slide of template.content.slides) {
        expect(slide.title.trim().length).toBeGreaterThan(0)
      }
    }
  })

  it('derives a non-empty structure preview for every template', () => {
    for (const template of TEMPLATES) {
      expect(templatePreview(template).length).toBeGreaterThan(0)
    }
  })

  it('maps every category to a distinct i18n key', () => {
    const keys = TEMPLATE_CATEGORIES.map(templateCategoryKey)
    expect(new Set(keys).size).toBe(keys.length)
    for (const key of keys) {
      expect(key).toMatch(/^templateCat/)
    }
  })
})

describe('template visual previews', () => {
  it('produces a non-empty visual preview model for every template', () => {
    for (const template of TEMPLATES) {
      const visual = templateVisualPreview(template)
      expect(visual.app).toBe(template.content.app)
      if (visual.app === 'docs') {
        expect(visual.blocks.length).toBeGreaterThan(0)
        const text = visual.blocks
          .map((block) => {
            switch (block.kind) {
              case 'heading':
              case 'paragraph':
                return block.text
              case 'list':
                return block.items.join(' ')
              case 'table':
                return block.rows
                  .flat()
                  .map((cell) => cell.text)
                  .join(' ')
            }
          })
          .join(' ')
        expect(text.trim().length).toBeGreaterThan(0)
      } else if (visual.app === 'sheets') {
        expect(visual.rows.length).toBeGreaterThan(0)
        expect(visual.rows[0]!.length).toBeGreaterThan(0)
        expect(visual.rows.flat().some((cell) => cell.text.length > 0)).toBe(true)
      } else {
        expect(visual.title.trim().length).toBeGreaterThan(0)
      }
    }
  })

  it('builds the docs preview from real headings, paragraphs, and tables', () => {
    const resume = TEMPLATES.find((template) => template.id === 'resume-modern')!
    const visual = templateVisualPreview(resume)
    expect(visual.app).toBe('docs')
    if (visual.app !== 'docs') return
    expect(visual.blocks[0]).toEqual({ kind: 'heading', level: 1, text: '[Your Full Name]' })
    expect(visual.blocks.some((block) => block.kind === 'paragraph')).toBe(true)

    const agenda = TEMPLATES.find((template) => template.id === 'meeting-agenda')!
    const agendaVisual = templateVisualPreview(agenda)
    expect(agendaVisual.app).toBe('docs')
    if (agendaVisual.app !== 'docs') return
    const table = agendaVisual.blocks.find((block) => block.kind === 'table')
    expect(table).toBeDefined()
    if (table?.kind !== 'table') return
    expect(table.rows[0]!.map((cell) => cell.text)).toEqual(['Time', 'Topic', 'Lead'])
    expect(table.rows[0]!.every((cell) => cell.bold)).toBe(true)
  })

  it('shows cached formula results in the sheets preview grid', () => {
    const template = TEMPLATES.find((candidate) => candidate.id === 'invoice-service')!
    const visual = templateVisualPreview(template)
    expect(visual.app).toBe('sheets')
    if (visual.app !== 'sheets') return
    expect(visual.sheetName).toBe('Invoice')
    expect(visual.headerRow).toBe(true)
    expect(visual.rows[0]!.map((cell) => cell.text)).toEqual([
      'Description',
      'Quantity',
      'Rate',
      'Amount',
    ])
    const allText = visual.rows.flat().map((cell) => cell.text)
    expect(allText).toContain('$500.00')
    expect(allText).toContain('$2,050.00')
    expect(allText).toContain('$2,193.50')
    // the bold total row keeps its bold flag for the grid styling
    const totalCell = visual.rows.flat().find((cell) => cell.text === '$2,193.50')!
    expect(totalCell.bold).toBe(true)
    expect(totalCell.numeric).toBe(true)
  })

  it('uses the first slide and counts the rest for the slides preview', () => {
    const template = TEMPLATES.find((candidate) => candidate.id === 'pitch-startup')!
    const visual = templateVisualPreview(template)
    expect(visual.app).toBe('slides')
    if (visual.app !== 'slides') return
    expect(visual.title).toBe('[Company Name]')
    expect(visual.subtitle).toContain('One-line pitch')
    expect(visual.extraSlides).toBe(7)
  })
})

describe('team-format templates (onboarding pack shapes, genericized)', () => {
  const EXPECTED: Array<{ id: string; category: string; app: string }> = [
    { id: 'team-working-preferences', category: 'meetings', app: 'docs' },
    { id: 'executive-presentation', category: 'presentations', app: 'slides' },
    { id: 'one-page-project-update', category: 'presentations', app: 'slides' },
    { id: 'financial-review', category: 'presentations', app: 'slides' },
    { id: 'process-timeline', category: 'presentations', app: 'slides' },
    { id: 'formatted-data-table', category: 'budgets', app: 'sheets' },
  ]

  it('includes all six templates in the right category and app', () => {
    for (const expected of EXPECTED) {
      const template = TEMPLATES.find((candidate) => candidate.id === expected.id)
      expect(template).toBeDefined()
      expect(template!.category).toBe(expected.category)
      expect(template!.content.app).toBe(expected.app)
    }
  })

  it('carries no source names, brands, banks, or real figures', () => {
    const forbidden = ['tyler', 'full sail', 'vilano', 'goldman', '439', '440']
    for (const expected of EXPECTED) {
      const template = TEMPLATES.find((candidate) => candidate.id === expected.id)!
      const text = JSON.stringify(template.content).toLowerCase()
      for (const word of forbidden) {
        expect(text).not.toContain(word)
      }
    }
  })

  it('builds the formatted data table with a header row and a correct total', async () => {
    const template = TEMPLATES.find((candidate) => candidate.id === 'formatted-data-table')!
    expect(template.content.app).toBe('sheets')
    if (template.content.app !== 'sheets') return
    const buffer = await buildTemplateXlsx(template.content.sheets)
    const zip = await JSZip.loadAsync(buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(sheet).toContain('<f>SUM(C2:C4)</f>')
    expect(sheet).toContain('<v>450</v>')
    expect(sheet).toContain('s="2"')
  })
})

describe('template file builders', () => {
  it('builds a multi-sheet workbook with formulas and header styling', async () => {
    const template = TEMPLATES.find((candidate) => candidate.id === 'invoice-service')!
    expect(template.content.app).toBe('sheets')
    if (template.content.app !== 'sheets') return
    const buffer = await buildTemplateXlsx(template.content.sheets)
    const zip = await JSZip.loadAsync(buffer)
    const workbook = await zip.file('xl/workbook.xml')!.async('text')
    expect(workbook).toContain('name="Invoice"')
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(sheet).toContain('<f>SUM(D2:D4)</f>')
    expect(sheet).toContain('<v>2050</v>')
    // header row cells carry the header style index
    expect(sheet).toContain('s="2"')
    const styles = await zip.file('xl/styles.xml')!.async('text')
    expect(styles).toContain('1C5B4A')
  })

  it('builds a deck with one slide per definition, in order, titles inside', async () => {
    const template = TEMPLATES.find((candidate) => candidate.id === 'pitch-startup')!
    expect(template.content.app).toBe('slides')
    if (template.content.app !== 'slides') return
    const bytes = await buildTemplatePptx(template.content.slides)
    const zip = await JSZip.loadAsync(bytes)
    const slideNames = Object.keys(zip.files).filter((name) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(name),
    )
    expect(slideNames.length).toBe(template.content.slides.length)
    // slide files are numbered in creation order, so each must hold its own title
    for (let i = 0; i < template.content.slides.length; i++) {
      const xml = await zip.file(`ppt/slides/slide${i + 1}.xml`)!.async('text')
      expect(xml).toContain(escapeXml(template.content.slides[i]!.title))
    }
  })
})
