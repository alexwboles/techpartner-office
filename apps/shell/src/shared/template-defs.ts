/**
 * Template gallery definitions (pure data — no node or Electron APIs, so
 * both the shell renderer gallery and the shell main builders can import
 * this module). All template content is written originally for TechPartner
 * Office; see docs/template-gallery-research.md for the competitor
 * categories this library was compared against.
 *
 * Docs templates are HTML fed through the app's own AI-document pipeline
 * (parseHtmlFragment): headings, paragraphs, lists and tables all work.
 * Use [bracketed placeholders] for text the user replaces, never fake
 * personal data.
 *
 * Sheets templates are row definitions. A cell is a plain string/number, or
 * a formula cell { formula, cached } where `cached` is the value a reader
 * shows before recalculation. Mark the first row a header with headerRow.
 * Widths are column widths in characters.
 *
 * Slides templates are slide definitions: the first slide is the title
 * slide (title + subtitle), the rest are title + bullets.
 */

export type TemplateApp = 'docs' | 'sheets' | 'slides'

export type TemplateCategory =
  | 'resumes'
  | 'cover-letters'
  | 'business-letters'
  | 'invoices-quotes'
  | 'business-plans'
  | 'meetings'
  | 'reports'
  | 'budgets'
  | 'schedules'
  | 'presentations'
  | 'flyers'
  | 'forms'
  | 'education'

export const TEMPLATE_CATEGORIES: readonly TemplateCategory[] = [
  'resumes',
  'cover-letters',
  'business-letters',
  'invoices-quotes',
  'business-plans',
  'meetings',
  'reports',
  'budgets',
  'schedules',
  'presentations',
  'flyers',
  'forms',
  'education',
]

/** i18n key for a category label (defined in strings.ts, all languages). */
export function templateCategoryKey(category: TemplateCategory): string {
  switch (category) {
    case 'resumes':
      return 'templateCatResumes'
    case 'cover-letters':
      return 'templateCatCoverLetters'
    case 'business-letters':
      return 'templateCatBusinessLetters'
    case 'invoices-quotes':
      return 'templateCatInvoices'
    case 'business-plans':
      return 'templateCatBusinessPlans'
    case 'meetings':
      return 'templateCatMeetings'
    case 'reports':
      return 'templateCatReports'
    case 'budgets':
      return 'templateCatBudgets'
    case 'schedules':
      return 'templateCatSchedules'
    case 'presentations':
      return 'templateCatPresentations'
    case 'flyers':
      return 'templateCatFlyers'
    case 'forms':
      return 'templateCatForms'
    case 'education':
      return 'templateCatEducation'
  }
}

export interface TemplateFormulaCell {
  formula: string
  cached: number | string
  format?: 'currency' | 'percent' | 'int'
  bold?: boolean
}

export type TemplateSheetCell = string | number | TemplateFormulaCell

export function isFormulaCell(cell: TemplateSheetCell): cell is TemplateFormulaCell {
  return typeof cell === 'object' && cell !== null && 'formula' in cell
}

export interface TemplateSheet {
  name: string
  headerRow?: boolean
  widths?: number[]
  rows: TemplateSheetCell[][]
}

export interface TemplateSlide {
  title: string
  subtitle?: string
  bullets: string[]
}

export type TemplateContent =
  | { app: 'docs'; html: string }
  | { app: 'sheets'; sheets: TemplateSheet[] }
  | { app: 'slides'; slides: TemplateSlide[] }

export interface TemplateDef {
  id: string
  title: string
  description: string
  category: TemplateCategory
  content: TemplateContent
}

export function templateAppOf(template: TemplateDef): TemplateApp {
  return template.content.app
}

/** file extension a template opens as */
export function templateExtOf(template: TemplateDef): 'docx' | 'xlsx' | 'pptx' {
  switch (template.content.app) {
    case 'docs':
      return 'docx'
    case 'sheets':
      return 'xlsx'
    case 'slides':
      return 'pptx'
  }
}

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
}

function docHeadings(html: string): string[] {
  const headings: string[] = []
  const re = /<h[1-6][^>]*>([\s\S]*?)<\/h[1-6]>/gi
  let match: RegExpExecArray | null
  while ((match = re.exec(html)) !== null) {
    const text = stripHtml(match[1] ?? '')
    if (text) headings.push(text)
  }
  return headings
}

/**
 * Structure preview lines for a template card / detail pane: headings for
 * docs, sheet names + row counts for sheets, slide titles for slides.
 * Never a fake screenshot — always derived from the real content.
 */
export function templatePreview(template: TemplateDef): string[] {
  switch (template.content.app) {
    case 'docs': {
      const headings = docHeadings(template.content.html)
      if (headings.length > 0) return headings.slice(0, 6)
      const text = stripHtml(template.content.html)
      return text ? [text.slice(0, 120)] : []
    }
    case 'sheets':
      return template.content.sheets.map((sheet) => `${sheet.name} — ${sheet.rows.length} rows`)
    case 'slides':
      return template.content.slides.map((slide) => slide.title)
  }
}

/** One cell of a visual preview grid (sheet rows, doc table rows). */
export interface TemplatePreviewCell {
  text: string
  bold: boolean
  numeric: boolean
}

/** One content block of a docs visual preview, in document order. */
export type TemplatePreviewBlock =
  | { kind: 'heading'; level: 1 | 2 | 3; text: string }
  | { kind: 'paragraph'; text: string }
  | { kind: 'list'; items: string[] }
  | { kind: 'table'; rows: TemplatePreviewCell[][] }

/**
 * Render-ready model for a card's miniature live preview. Like the docs
 * pipeline itself (parseHtmlFragment), the docs model is produced by a
 * whitelist parse of the template HTML into safe blocks — the gallery
 * renders these as React elements, never as injected HTML. Everything is
 * truncated to the first screenful so 58 cards stay cheap to render.
 */
export type TemplateVisualPreview =
  | { app: 'docs'; blocks: TemplatePreviewBlock[] }
  | {
      app: 'sheets'
      sheetName: string
      headerRow: boolean
      rows: TemplatePreviewCell[][]
      extraSheets: number
    }
  | { app: 'slides'; title: string; subtitle?: string; lines: string[]; extraSlides: number }

const PREVIEW_BLOCK_LIMIT = 9
const PREVIEW_LIST_LIMIT = 6
const PREVIEW_TABLE_ROW_LIMIT = 6
const PREVIEW_TABLE_COL_LIMIT = 5
const PREVIEW_SHEET_ROW_LIMIT = 7
const PREVIEW_SHEET_COL_LIMIT = 5

function docPreviewBlocks(html: string): TemplatePreviewBlock[] {
  const blocks: TemplatePreviewBlock[] = []
  const blockRe = /<(h[1-3]|p|ul|table)\b[^>]*>([\s\S]*?)<\/\1>/gi
  let match: RegExpExecArray | null
  while ((match = blockRe.exec(html)) !== null && blocks.length < PREVIEW_BLOCK_LIMIT) {
    const tag = match[1]!.toLowerCase()
    const inner = match[2]!
    if (tag === 'ul') {
      const items: string[] = []
      const liRe = /<li[^>]*>([\s\S]*?)<\/li>/gi
      let li: RegExpExecArray | null
      while ((li = liRe.exec(inner)) !== null && items.length < PREVIEW_LIST_LIMIT) {
        const text = stripHtml(li[1]!)
        if (text) items.push(text)
      }
      if (items.length > 0) blocks.push({ kind: 'list', items })
    } else if (tag === 'table') {
      const rows: TemplatePreviewCell[][] = []
      const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi
      let tr: RegExpExecArray | null
      while ((tr = trRe.exec(inner)) !== null && rows.length < PREVIEW_TABLE_ROW_LIMIT) {
        const cells: TemplatePreviewCell[] = []
        const cellRe = /<(th|td)\b[^>]*>([\s\S]*?)<\/\1>/gi
        let cell: RegExpExecArray | null
        while ((cell = cellRe.exec(tr[1]!)) !== null && cells.length < PREVIEW_TABLE_COL_LIMIT) {
          cells.push({
            text: stripHtml(cell[2]!),
            bold: cell[1]!.toLowerCase() === 'th',
            numeric: false,
          })
        }
        if (cells.length > 0) rows.push(cells)
      }
      if (rows.length > 0) blocks.push({ kind: 'table', rows })
    } else if (tag === 'p') {
      const text = stripHtml(inner)
      if (text) blocks.push({ kind: 'paragraph', text: text.slice(0, 300) })
    } else {
      const text = stripHtml(inner)
      if (text) {
        blocks.push({ kind: 'heading', level: Number(tag[1]) as 1 | 2 | 3, text })
      }
    }
  }
  return blocks
}

function formatPreviewNumber(value: number, format?: 'currency' | 'percent' | 'int'): string {
  if (format === 'currency') {
    return `$${value.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`
  }
  if (format === 'percent') {
    const pct = Math.abs(value) <= 1 ? value * 100 : value
    return `${Math.round(pct * 10) / 10}%`
  }
  if (format === 'int') return String(Math.round(value))
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100)
}

function sheetCellPreview(cell: TemplateSheetCell): TemplatePreviewCell {
  if (typeof cell === 'string') return { text: cell, bold: false, numeric: false }
  if (typeof cell === 'number') {
    return { text: formatPreviewNumber(cell), bold: false, numeric: true }
  }
  // Formula cells show their cached result, the way a reader sees the sheet
  // before recalculation.
  const cached = cell.cached
  return {
    text: typeof cached === 'number' ? formatPreviewNumber(cached, cell.format) : cached,
    bold: cell.bold === true,
    numeric: typeof cached === 'number',
  }
}

/**
 * Visual preview model for a gallery card: the first screenful of the
 * template's real content — doc blocks for docs, a grid of the first
 * sheet for sheets, the title slide for slides.
 */
export function templateVisualPreview(template: TemplateDef): TemplateVisualPreview {
  switch (template.content.app) {
    case 'docs': {
      const blocks = docPreviewBlocks(template.content.html)
      if (blocks.length > 0) return { app: 'docs', blocks }
      const text = stripHtml(template.content.html)
      return { app: 'docs', blocks: text ? [{ kind: 'paragraph', text: text.slice(0, 300) }] : [] }
    }
    case 'sheets': {
      const sheet = template.content.sheets[0]!
      return {
        app: 'sheets',
        sheetName: sheet.name,
        headerRow: sheet.headerRow === true,
        rows: sheet.rows
          .slice(0, PREVIEW_SHEET_ROW_LIMIT)
          .map((row) => row.slice(0, PREVIEW_SHEET_COL_LIMIT).map(sheetCellPreview)),
        extraSheets: template.content.sheets.length - 1,
      }
    }
    case 'slides': {
      const first = template.content.slides[0]!
      return {
        app: 'slides',
        title: first.title,
        ...(first.subtitle ? { subtitle: first.subtitle } : {}),
        lines: first.bullets.slice(0, 4),
        extraSlides: template.content.slides.length - 1,
      }
    }
  }
}

export const TEMPLATES: TemplateDef[] = [
  {
    id: 'resume-modern',
    title: 'Modern Resume',
    description:
      'Clean single-column resume with summary, experience, skills, and education sections. Replace the placeholders with your own history.',
    category: 'resumes',
    content: {
      app: 'docs',
      html: [
        '<h1>[Your Full Name]</h1>',
        '<p>[City, State] · [Phone] · [Email] · [LinkedIn or portfolio URL]</p>',
        '<h2>Summary</h2>',
        '<p>[Two or three sentences about what you do, your strongest skills, and the kind of role you are looking for.]</p>',
        '<h2>Experience</h2>',
        '<h3>[Job Title] — [Company], [City, State]</h3>',
        '<p>[Start Month Year] – [End Month Year or Present]</p>',
        '<ul><li>[Achievement with a number: what you did, and the result it produced.]</li><li>[A second achievement, same pattern.]</li><li>[A third achievement, same pattern.]</li></ul>',
        '<h3>[Previous Job Title] — [Company], [City, State]</h3>',
        '<p>[Start Month Year] – [End Month Year]</p>',
        '<ul><li>[Achievement with a number.]</li><li>[Achievement with a number.]</li></ul>',
        '<h2>Skills</h2>',
        '<p>[Skill 1] · [Skill 2] · [Skill 3] · [Skill 4] · [Skill 5]</p>',
        '<h2>Education</h2>',
        '<p>[Degree], [School] — [Graduation Year]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'invoice-service',
    title: 'Service Invoice',
    description:
      'Invoice with line items, tax, and a total that recalculates from the rates and quantities you enter.',
    category: 'invoices-quotes',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Invoice',
          headerRow: true,
          widths: [42, 12, 14, 16],
          rows: [
            ['Description', 'Quantity', 'Rate', 'Amount'],
            [
              'Consulting — project discovery',
              4,
              125,
              { formula: 'B2*C2', cached: 500, format: 'currency' },
            ],
            [
              'Consulting — implementation',
              10,
              125,
              { formula: 'B3*C3', cached: 1250, format: 'currency' },
            ],
            [
              'Monthly support retainer',
              1,
              300,
              { formula: 'B4*C4', cached: 300, format: 'currency' },
            ],
            [
              '',
              '',
              'Subtotal',
              { formula: 'SUM(D2:D4)', cached: 2050, format: 'currency', bold: true },
            ],
            ['', '', 'Tax (7%)', { formula: 'D5*0.07', cached: 143.5, format: 'currency' }],
            [
              '',
              '',
              'Total due',
              { formula: 'D5+D6', cached: 2193.5, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'pitch-startup',
    title: 'Startup Pitch Deck',
    description:
      'Eight-slide investor pitch: problem, solution, market, product, business model, traction, team, and the ask.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Company Name]',
          subtitle: '[One-line pitch: who you serve and the outcome you deliver]',
          bullets: [],
        },
        {
          title: 'The Problem',
          bullets: [
            '[Who feels this pain, and how often?]',
            '[What does it cost them today — in money, time, or risk?]',
            '[Why existing options fall short.]',
          ],
        },
        {
          title: 'The Solution',
          bullets: [
            '[Your product in one sentence.]',
            '[The three things it does differently.]',
            '[What a customer experiences in the first week.]',
          ],
        },
        {
          title: 'Market Opportunity',
          bullets: [
            '[Total market, serviceable market, and your first beachhead.]',
            '[Why now: the shift that makes this winnable today.]',
          ],
        },
        {
          title: 'Product',
          bullets: ['[How it works, in three steps.]', '[What is live today vs. on the roadmap.]'],
        },
        {
          title: 'Business Model',
          bullets: [
            '[How you charge, and what a typical customer pays.]',
            '[Unit economics: cost to acquire vs. lifetime value.]',
          ],
        },
        {
          title: 'Traction',
          bullets: [
            '[Customers, revenue, or pilots — with numbers.]',
            '[Growth over the last three months.]',
          ],
        },
        {
          title: 'Team & The Ask',
          bullets: [
            '[Founders and why they are the right people for this.]',
            '[How much you are raising and what it funds.]',
            '[The milestone the raise gets you to.]',
          ],
        },
      ],
    },
  },
  {
    id: 'resume-chronological',
    title: 'Chronological Resume',
    description:
      'Classic reverse-chronological resume that leads with your most recent role and walks back through your work history in order.',
    category: 'resumes',
    content: {
      app: 'docs',
      html: [
        '<h1>[Your Full Name]</h1>',
        '<p>[City, State] · [Phone] · [Email]</p>',
        '<h2>Professional Summary</h2>',
        '<p>[One or two sentences: your profession, years of experience, and the value you bring to an employer.]</p>',
        '<h2>Work Experience</h2>',
        '<h3>[Most Recent Job Title] — [Company], [City, State]</h3>',
        '<p>[Start Month Year] – [End Month Year or Present]</p>',
        '<ul><li>[Responsibility or achievement, with a measurable result where possible.]</li><li>[Responsibility or achievement.]</li><li>[Responsibility or achievement.]</li></ul>',
        '<h3>[Earlier Job Title] — [Company], [City, State]</h3>',
        '<p>[Start Month Year] – [End Month Year]</p>',
        '<ul><li>[Responsibility or achievement.]</li><li>[Responsibility or achievement.]</li></ul>',
        '<h3>[Earliest Job Title] — [Company], [City, State]</h3>',
        '<p>[Start Month Year] – [End Month Year]</p>',
        '<ul><li>[Responsibility or achievement.]</li></ul>',
        '<h2>Education</h2>',
        '<p>[Degree], [School] — [Graduation Year]</p>',
        '<h2>Certifications</h2>',
        '<p>[Certification], [Issuing Body] — [Year]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'resume-student',
    title: 'Student Resume',
    description:
      'Resume for students and recent graduates that highlights education, coursework, projects, and part-time or volunteer experience.',
    category: 'resumes',
    content: {
      app: 'docs',
      html: [
        '<h1>[Your Full Name]</h1>',
        '<p>[City, State] · [Phone] · [Email]</p>',
        '<h2>Objective</h2>',
        '<p>[What you are studying, what kind of role or internship you are seeking, and what you bring to it.]</p>',
        '<h2>Education</h2>',
        '<p>[Degree or Program], [School] — Expected [Graduation Year]</p>',
        '<p>Relevant coursework: [Course 1], [Course 2], [Course 3]</p>',
        '<h2>Projects</h2>',
        '<h3>[Project Name]</h3>',
        '<p>[What you built or researched, the tools you used, and what it achieved.]</p>',
        '<h3>[Project Name]</h3>',
        '<p>[What you built or researched, the tools you used, and what it achieved.]</p>',
        '<h2>Experience</h2>',
        '<h3>[Role] — [Employer or Organization]</h3>',
        '<p>[Start Month Year] – [End Month Year or Present]</p>',
        '<ul><li>[What you did and what came of it.]</li><li>[What you did and what came of it.]</li></ul>',
        '<h2>Volunteer Work</h2>',
        '<p>[Role], [Organization] — [Dates]: [Brief description of your contribution.]</p>',
        '<h2>Skills</h2>',
        '<p>[Skill 1] · [Skill 2] · [Skill 3] · [Skill 4]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'resume-executive',
    title: 'Executive Resume',
    description:
      'Senior-leader resume built around a leadership profile, scope of responsibility, and headline business outcomes.',
    category: 'resumes',
    content: {
      app: 'docs',
      html: [
        '<h1>[Your Full Name]</h1>',
        '<p>[City, State] · [Phone] · [Email] · [LinkedIn URL]</p>',
        '<h2>Leadership Profile</h2>',
        '<p>[Two or three sentences: the scale of organizations you have led, your core leadership strengths, and the results you are known for.]</p>',
        '<h2>Career Highlights</h2>',
        '<ul><li>[Headline outcome: revenue growth, turnaround, expansion, or cost saving, with a number.]</li><li>[Headline outcome with a number.]</li><li>[Headline outcome with a number.]</li></ul>',
        '<h2>Experience</h2>',
        '<h3>[Title] — [Company], [City, State]</h3>',
        '<p>[Start Year] – [End Year or Present] · Scope: [Team size, budget, or regions led]</p>',
        '<ul><li>[Key initiative and its business result.]</li><li>[Key initiative and its business result.]</li></ul>',
        '<h3>[Earlier Title] — [Company], [City, State]</h3>',
        '<p>[Start Year] – [End Year] · Scope: [Team size, budget, or regions led]</p>',
        '<ul><li>[Key initiative and its business result.]</li></ul>',
        '<h2>Education</h2>',
        '<p>[Degree], [School] — [Graduation Year]</p>',
        '<h2>Board and Advisory Roles</h2>',
        '<p>[Organization], [Role] — [Dates]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'cover-letter-general',
    title: 'General Cover Letter',
    description:
      'Standard application cover letter: the role, why you fit it, your strongest evidence, and a clear close.',
    category: 'cover-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[City, State] · [Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>[Hiring Manager Name or Hiring Team]<br>[Company]<br>[Company Address]</p>',
        '<h2>Application for [Job Title] at [Company]</h2>',
        '<p>Dear [Hiring Manager Name or Hiring Team],</p>',
        '<p>I am applying for the [Job Title] position at [Company]. [One sentence on why this role and this company appeal to you.]</p>',
        '<p>In my current role as [Current Role] at [Current Employer], I [strongest relevant achievement, with a number where possible]. I also bring [second strength or skill] and [third strength or skill], which match the requirements listed for this position.</p>',
        '<p>What draws me to [Company] is [something specific: its work, values, product, or reputation]. I would welcome the chance to bring [your key contribution] to your team.</p>',
        '<p>Thank you for considering my application. I would be glad to discuss how I can contribute at your convenience.</p>',
        '<p>Sincerely,<br>[Your Full Name]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'cover-letter-career-change',
    title: 'Career Change Cover Letter',
    description:
      'Cover letter for moving into a new field, framing transferable skills and the reason for the change up front.',
    category: 'cover-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[City, State] · [Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>[Hiring Manager Name or Hiring Team]<br>[Company]</p>',
        '<h2>Application for [Job Title] at [Company]</h2>',
        '<p>Dear [Hiring Manager Name or Hiring Team],</p>',
        '<p>I am applying for the [Job Title] position at [Company] as I move from [Current Field] into [Target Field]. [One sentence on what sparked the change.]</p>',
        '<p>Although my background is in [Current Field], the skills transfer directly: [transferable skill 1] from [example], [transferable skill 2] from [example], and [transferable skill 3] from [example]. To prepare for this move, I have [course, certification, project, or volunteer work completed].</p>',
        '<p>I am drawn to [Company] because [specific reason], and I am ready to contribute [what you offer] from day one while I continue building [skill you are developing].</p>',
        '<p>Thank you for your consideration. I would welcome the opportunity to discuss my background and this transition with you.</p>',
        '<p>Sincerely,<br>[Your Full Name]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'cover-letter-internship',
    title: 'Internship Cover Letter',
    description:
      'Cover letter for students applying for internships, built on coursework, projects, and enthusiasm for the field.',
    category: 'cover-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[City, State] · [Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>[Hiring Manager Name or Hiring Team]<br>[Company]</p>',
        '<h2>Application for the [Internship Title] Internship</h2>',
        '<p>Dear [Hiring Manager Name or Hiring Team],</p>',
        '<p>I am a [Year, e.g. third-year] student studying [Major] at [School], and I am applying for the [Internship Title] internship at [Company] for [Term or Season].</p>',
        '<p>Through my coursework in [Course 1] and [Course 2], and my project on [Project Description], I have developed [Skill 1] and [Skill 2]. In [part-time job, club, or volunteer role], I [relevant achievement or responsibility].</p>',
        '<p>I am excited about [Company] because [specific reason: its work, mission, or product]. This internship would let me apply what I have learned while contributing [what you can offer the team].</p>',
        '<p>Thank you for considering my application. I would be glad to discuss my background and availability.</p>',
        '<p>Sincerely,<br>[Your Full Name]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'letter-formal-inquiry',
    title: 'Formal Inquiry Letter',
    description:
      'Polite business letter for requesting information about a product, service, partnership, or opportunity.',
    category: 'business-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[Your Company]<br>[Your Address]<br>[Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>[Recipient Name]<br>[Recipient Company]<br>[Recipient Address]</p>',
        '<h2>Inquiry Regarding [Product, Service, or Opportunity]</h2>',
        '<p>Dear [Recipient Name],</p>',
        '<p>I am writing to inquire about [product, service, or opportunity]. [One sentence on who you are and why you are asking.]</p>',
        '<p>Specifically, I would like to know:</p>',
        '<ul><li>[Question 1]</li><li>[Question 2]</li><li>[Question 3]</li></ul>',
        '<p>Any brochures, pricing details, or timelines you can share would be appreciated. Please let me know if a brief call would be easier.</p>',
        '<p>Thank you for your time. I look forward to hearing from you.</p>',
        '<p>Sincerely,<br>[Your Full Name]<br>[Your Title]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'letter-business-thank-you',
    title: 'Business Thank-You Letter',
    description:
      'Professional thank-you letter for a client, partner, or colleague after a meeting, deal, or act of support.',
    category: 'business-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[Your Company]</p>',
        '<p>[Date]</p>',
        '<h2>Thank You</h2>',
        '<p>Dear [Recipient Name],</p>',
        '<p>Thank you for [meeting with us, your business, your support, or your partnership] on [Date or Occasion].</p>',
        '<p>[One or two sentences on what you appreciated most: their time, insight, trust, or contribution.]</p>',
        '<p>I especially valued [specific detail from the meeting or project]. As a next step, we will [agreed action] by [Date].</p>',
        '<p>I look forward to working with you again and to [shared goal or future collaboration].</p>',
        '<p>With appreciation,<br>[Your Full Name]<br>[Your Title]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'letter-complaint-resolution',
    title: 'Complaint and Issue Resolution Letter',
    description:
      'Firm but courteous complaint letter that states the problem, the evidence, and the resolution you are asking for.',
    category: 'business-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Your Full Name]<br>[Your Address]<br>[Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>[Recipient Name or Customer Service Department]<br>[Company]<br>[Company Address]</p>',
        '<h2>Complaint Regarding [Product or Service]</h2>',
        '<p>Dear [Recipient Name or Customer Service Team],</p>',
        '<p>I am writing about [product or service] purchased on [Date], [order or account reference]. [Clear one-sentence description of the problem.]</p>',
        '<p>What happened: [Factual account of events in chronological order, including dates and any previous contacts.]</p>',
        '<p>To resolve this, I am asking for [refund, replacement, repair, or other specific resolution]. I have enclosed [receipts, photos, or correspondence] as supporting evidence.</p>',
        '<p>Please respond by [Date] so we can settle this promptly. I hope we can resolve the matter without further escalation.</p>',
        '<p>Sincerely,<br>[Your Full Name]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'letter-job-offer',
    title: 'Job Offer Letter',
    description:
      'Formal offer letter covering role, start date, compensation, benefits, and the acceptance deadline.',
    category: 'business-letters',
    content: {
      app: 'docs',
      html: [
        '<p>[Company Name]<br>[Company Address]<br>[Phone] · [Email]</p>',
        '<p>[Date]</p>',
        '<p>Dear [Candidate Name],</p>',
        '<p>We are pleased to offer you the position of [Job Title] at [Company Name], reporting to [Manager Name]. Your anticipated start date is [Start Date].</p>',
        '<h2>Compensation and Benefits</h2>',
        '<ul><li>Base salary: [Annual Salary] per year, paid [pay schedule].</li><li>Work schedule: [Hours and location, including remote or hybrid details].</li><li>Benefits: [Health coverage, paid time off, retirement plan, and other benefits].</li><li>[Any bonus, equity, or additional terms].</li></ul>',
        '<p>This offer is contingent on [background check, references, or other conditions, if any]. Please confirm your acceptance by signing below and returning this letter by [Acceptance Deadline].</p>',
        '<p>We look forward to welcoming you to the team.</p>',
        '<p>Sincerely,<br>[Hiring Manager Name]<br>[Title], [Company Name]</p>',
        '<p>Accepted by: ______________________ Date: __________</p>',
      ].join('\n'),
    },
  },
  {
    id: 'invoice-product-sales',
    title: 'Product Sales Invoice',
    description:
      'Invoice for selling physical products, with quantities, unit prices, tax, and a total that recalculates automatically.',
    category: 'invoices-quotes',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Invoice',
          headerRow: true,
          widths: [42, 12, 14, 16],
          rows: [
            ['Product', 'Quantity', 'Unit Price', 'Amount'],
            ['[Product 1]', 2, 45, { formula: 'B2*C2', cached: 90, format: 'currency' }],
            ['[Product 2]', 3, 30, { formula: 'B3*C3', cached: 90, format: 'currency' }],
            ['[Product 3]', 1, 120, { formula: 'B4*C4', cached: 120, format: 'currency' }],
            [
              '',
              '',
              'Subtotal',
              { formula: 'SUM(D2:D4)', cached: 300, format: 'currency', bold: true },
            ],
            ['', '', 'Tax (7%)', { formula: 'D5*0.07', cached: 21, format: 'currency' }],
            [
              '',
              '',
              'Total due',
              { formula: 'D5+D6', cached: 321, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'quote-estimate',
    title: 'Quote and Estimate',
    description:
      'Client-facing quote with line items, tax, and a total, plus a validity note so the client knows how long the price holds.',
    category: 'invoices-quotes',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Quote',
          headerRow: true,
          widths: [42, 12, 14, 16],
          rows: [
            ['Item', 'Quantity', 'Unit Price', 'Amount'],
            ['[Service or item 1]', 5, 80, { formula: 'B2*C2', cached: 400, format: 'currency' }],
            ['[Service or item 2]', 2, 150, { formula: 'B3*C3', cached: 300, format: 'currency' }],
            ['[Service or item 3]', 1, 250, { formula: 'B4*C4', cached: 250, format: 'currency' }],
            [
              '',
              '',
              'Subtotal',
              { formula: 'SUM(D2:D4)', cached: 950, format: 'currency', bold: true },
            ],
            ['', '', 'Tax (7%)', { formula: 'D5*0.07', cached: 66.5, format: 'currency' }],
            ['', '', 'Total', { formula: 'D5+D6', cached: 1016.5, format: 'currency', bold: true }],
            ['Valid until', '[Date]', '', ''],
          ],
        },
      ],
    },
  },
  {
    id: 'payment-receipt',
    title: 'Payment Receipt',
    description:
      'Receipt showing items purchased, tax, the total, the amount paid, and a zero balance when payment is complete.',
    category: 'invoices-quotes',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Receipt',
          headerRow: true,
          widths: [42, 12, 14, 16],
          rows: [
            ['Item', 'Quantity', 'Unit Price', 'Amount'],
            ['[Item 1]', 1, 199, { formula: 'B2*C2', cached: 199, format: 'currency' }],
            ['[Item 2]', 2, 49, { formula: 'B3*C3', cached: 98, format: 'currency' }],
            [
              '',
              '',
              'Subtotal',
              { formula: 'SUM(D2:D3)', cached: 297, format: 'currency', bold: true },
            ],
            ['', '', 'Tax (7%)', { formula: 'D4*0.07', cached: 20.79, format: 'currency' }],
            ['', '', 'Total', { formula: 'D4+D5', cached: 317.79, format: 'currency', bold: true }],
            ['', '', 'Amount paid', 317.79],
            [
              '',
              '',
              'Balance due',
              { formula: 'D6-D7', cached: 0, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'plan-lean-business',
    title: 'Lean Business Plan',
    description:
      'One-page lean business plan covering the problem, solution, customers, channels, costs, and revenue in brief sections.',
    category: 'business-plans',
    content: {
      app: 'docs',
      html: [
        '<h1>Lean Business Plan: [Business Name]</h1>',
        '<p>[One-sentence description of the business.]</p>',
        '<h2>Problem</h2>',
        '<p>[The problem your customers face, and who faces it.]</p>',
        '<h2>Solution</h2>',
        '<p>[Your product or service, and how it solves the problem.]</p>',
        '<h2>Target Customers</h2>',
        '<p>[Who your first customers are, and how you will reach them.]</p>',
        '<h2>Unique Value</h2>',
        '<p>[What makes your offer different from existing options.]</p>',
        '<h2>Channels</h2>',
        '<p>[How customers discover and buy from you.]</p>',
        '<h2>Revenue</h2>',
        '<p>[How you charge, and what a typical customer pays.]</p>',
        '<h2>Costs</h2>',
        '<p>[Your main fixed and variable costs.]</p>',
        '<h2>Key Metrics</h2>',
        '<p>[The numbers you will watch to know the business is working.]</p>',
        '<h2>Next Steps</h2>',
        '<ul><li>[Immediate action 1]</li><li>[Immediate action 2]</li><li>[Immediate action 3]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'proposal-services',
    title: 'Services Proposal',
    description:
      'Proposal for selling a service: scope, deliverables, timeline, pricing, and terms the client can accept.',
    category: 'business-plans',
    content: {
      app: 'docs',
      html: [
        '<h1>Services Proposal: [Service Name]</h1>',
        '<p>Prepared for [Client Name] by [Your Company] on [Date].</p>',
        '<h2>Overview</h2>',
        '<p>[What the client needs, and how your service meets that need.]</p>',
        '<h2>Scope of Work</h2>',
        '<ul><li>[Deliverable or task 1]</li><li>[Deliverable or task 2]</li><li>[Deliverable or task 3]</li></ul>',
        '<p>Out of scope: [Work explicitly not included.]</p>',
        '<h2>Timeline</h2>',
        '<p>[Phase 1 and dates], [Phase 2 and dates], [Phase 3 and dates].</p>',
        '<h2>Pricing</h2>',
        '<table><tr><th>Item</th><th>Price</th></tr><tr><td>[Service component 1]</td><td>[Amount]</td></tr><tr><td>[Service component 2]</td><td>[Amount]</td></tr><tr><td>Total</td><td>[Total amount]</td></tr></table>',
        '<h2>Terms</h2>',
        '<p>[Payment schedule, revision policy, and cancellation terms.]</p>',
        '<h2>Acceptance</h2>',
        '<p>Accepted by: ______________________ Date: __________</p>',
      ].join('\n'),
    },
  },
  {
    id: 'proposal-project',
    title: 'Project Proposal',
    description:
      'Project proposal with background, objectives, approach, schedule, budget, and risks for approval by a sponsor.',
    category: 'business-plans',
    content: {
      app: 'docs',
      html: [
        '<h1>Project Proposal: [Project Name]</h1>',
        '<p>Prepared by [Your Name or Team] for [Sponsor or Client] on [Date].</p>',
        '<h2>Background</h2>',
        '<p>[Why this project is needed, and what situation it responds to.]</p>',
        '<h2>Objectives</h2>',
        '<ul><li>[Measurable objective 1]</li><li>[Measurable objective 2]</li><li>[Measurable objective 3]</li></ul>',
        '<h2>Approach</h2>',
        '<p>[How the work will be done, in phases or stages.]</p>',
        '<h2>Schedule</h2>',
        '<table><tr><th>Phase</th><th>Dates</th><th>Deliverable</th></tr><tr><td>[Phase 1]</td><td>[Dates]</td><td>[Deliverable]</td></tr><tr><td>[Phase 2]</td><td>[Dates]</td><td>[Deliverable]</td></tr></table>',
        '<h2>Budget</h2>',
        '<p>[Total budget and its main components.]</p>',
        '<h2>Risks</h2>',
        '<ul><li>[Risk 1 and how it will be handled]</li><li>[Risk 2 and how it will be handled]</li></ul>',
        '<h2>Approval</h2>',
        '<p>Approved by: ______________________ Date: __________</p>',
      ].join('\n'),
    },
  },
  {
    id: 'proposal-partnership',
    title: 'Partnership Proposal',
    description:
      'Proposal for a business partnership: shared goals, what each side brings, how it works, and how value is shared.',
    category: 'business-plans',
    content: {
      app: 'docs',
      html: [
        '<h1>Partnership Proposal: [Your Company] and [Partner Company]</h1>',
        '<p>Prepared on [Date] for [Partner Contact Name].</p>',
        '<h2>Why Partner</h2>',
        '<p>[The opportunity you see together, and why each side benefits.]</p>',
        '<h2>Shared Goals</h2>',
        '<ul><li>[Goal 1]</li><li>[Goal 2]</li><li>[Goal 3]</li></ul>',
        '<h2>What Each Side Brings</h2>',
        '<table><tr><th>Partner</th><th>Contribution</th></tr><tr><td>[Your Company]</td><td>[Resources, audience, or expertise]</td></tr><tr><td>[Partner Company]</td><td>[Resources, audience, or expertise]</td></tr></table>',
        '<h2>How It Works</h2>',
        '<p>[Day-to-day structure: who does what, and how decisions are made.]</p>',
        '<h2>Sharing Value</h2>',
        '<p>[How revenue, costs, or leads are shared.]</p>',
        '<h2>Next Steps</h2>',
        '<ul><li>[Step 1]</li><li>[Step 2]</li><li>[Step 3]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'meeting-agenda',
    title: 'Meeting Agenda',
    description:
      'Structured meeting agenda with attendees, objectives, timed discussion topics, and space for decisions and next steps.',
    category: 'meetings',
    content: {
      app: 'docs',
      html: [
        '<h1>Meeting Agenda: [Meeting Title]</h1>',
        '<p>Date: [Date] · Time: [Start Time] – [End Time] · Location: [Room or Video Link]</p>',
        '<h2>Attendees</h2>',
        '<p>[Name 1], [Name 2], [Name 3]</p>',
        '<h2>Objective</h2>',
        '<p>[What this meeting must decide or produce by the end.]</p>',
        '<h2>Agenda</h2>',
        '<table><tr><th>Time</th><th>Topic</th><th>Lead</th></tr><tr><td>[Time]</td><td>[Topic 1]</td><td>[Name]</td></tr><tr><td>[Time]</td><td>[Topic 2]</td><td>[Name]</td></tr><tr><td>[Time]</td><td>[Topic 3]</td><td>[Name]</td></tr></table>',
        '<h2>Pre-Reading</h2>',
        '<p>[Documents or links attendees should review beforehand.]</p>',
        '<h2>Decisions Needed</h2>',
        '<ul><li>[Decision 1]</li><li>[Decision 2]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'meeting-minutes',
    title: 'Meeting Minutes',
    description:
      'Formal meeting minutes recording attendees, discussion points, decisions made, and assigned action items.',
    category: 'meetings',
    content: {
      app: 'docs',
      html: [
        '<h1>Meeting Minutes: [Meeting Title]</h1>',
        '<p>Date: [Date] · Time: [Start Time] – [End Time] · Location: [Room or Video Link]</p>',
        '<h2>Attendees</h2>',
        '<p>Present: [Name 1], [Name 2], [Name 3]<br>Absent: [Name]</p>',
        '<h2>Discussion</h2>',
        '<h3>[Topic 1]</h3>',
        '<p>[Summary of what was discussed and the main points raised.]</p>',
        '<h3>[Topic 2]</h3>',
        '<p>[Summary of what was discussed and the main points raised.]</p>',
        '<h2>Decisions</h2>',
        '<ul><li>[Decision 1 and who made it]</li><li>[Decision 2 and who made it]</li></ul>',
        '<h2>Action Items</h2>',
        '<table><tr><th>Action</th><th>Owner</th><th>Due</th></tr><tr><td>[Action 1]</td><td>[Name]</td><td>[Date]</td></tr><tr><td>[Action 2]</td><td>[Name]</td><td>[Date]</td></tr></table>',
        '<h2>Next Meeting</h2>',
        '<p>[Date and time of the next meeting, if scheduled.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'meeting-weekly-team-notes',
    title: 'Weekly Team Notes',
    description:
      'Weekly team update note: wins, progress, blockers, and priorities for the coming week, in one shareable page.',
    category: 'meetings',
    content: {
      app: 'docs',
      html: [
        '<h1>Weekly Team Notes: Week of [Date]</h1>',
        '<p>Team: [Team Name] · Prepared by: [Name]</p>',
        '<h2>Wins This Week</h2>',
        '<ul><li>[Win 1]</li><li>[Win 2]</li></ul>',
        '<h2>Progress Updates</h2>',
        '<ul><li>[Name]: [What they worked on and where it stands]</li><li>[Name]: [What they worked on and where it stands]</li></ul>',
        '<h2>Blockers</h2>',
        '<ul><li>[Blocker and what is needed to clear it]</li></ul>',
        '<h2>Priorities for Next Week</h2>',
        '<ul><li>[Priority 1]</li><li>[Priority 2]</li><li>[Priority 3]</li></ul>',
        '<h2>Announcements</h2>',
        '<p>[Any dates, deadlines, or news the team should know.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'meeting-one-on-one-notes',
    title: 'One-on-One Notes',
    description:
      'One-on-one meeting notes for a manager and report: check-in, discussion topics, feedback, and agreed actions.',
    category: 'meetings',
    content: {
      app: 'docs',
      html: [
        '<h1>One-on-One: [Manager Name] and [Employee Name]</h1>',
        '<p>Date: [Date] · Cadence: [Weekly or Biweekly]</p>',
        '<h2>Check-In</h2>',
        '<p>[How things are going overall: workload, energy, and anything on their mind.]</p>',
        '<h2>Discussion Topics</h2>',
        '<ul><li>[Topic 1 and key points]</li><li>[Topic 2 and key points]</li></ul>',
        '<h2>Feedback</h2>',
        '<p>[Feedback shared in either direction, and what was agreed.]</p>',
        '<h2>Career and Growth</h2>',
        '<p>[Goals, development interests, or support the employee asked for.]</p>',
        '<h2>Action Items</h2>',
        '<ul><li>[Action] — Owner: [Name], Due: [Date]</li><li>[Action] — Owner: [Name], Due: [Date]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'report-project-status',
    title: 'Project Status Report',
    description:
      'Project status report with overall health, milestones, risks, and next steps for stakeholders and sponsors.',
    category: 'reports',
    content: {
      app: 'docs',
      html: [
        '<h1>Project Status Report: [Project Name]</h1>',
        '<p>Reporting period: [Start Date] – [End Date] · Prepared by: [Name]</p>',
        '<h2>Overall Status</h2>',
        '<p>Status: [On Track / At Risk / Off Track]<br>[One or two sentences explaining the status.]</p>',
        '<h2>Progress This Period</h2>',
        '<ul><li>[Completed work 1]</li><li>[Completed work 2]</li></ul>',
        '<h2>Milestones</h2>',
        '<table><tr><th>Milestone</th><th>Due</th><th>Status</th></tr><tr><td>[Milestone 1]</td><td>[Date]</td><td>[Complete / In Progress / Not Started]</td></tr><tr><td>[Milestone 2]</td><td>[Date]</td><td>[Complete / In Progress / Not Started]</td></tr></table>',
        '<h2>Risks and Issues</h2>',
        '<ul><li>[Risk or issue and the plan to address it]</li></ul>',
        '<h2>Next Steps</h2>',
        '<ul><li>[Next step 1]</li><li>[Next step 2]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'report-monthly-sales',
    title: 'Monthly Sales Report',
    description:
      'Monthly sales report summarizing revenue, top products, channel performance, and the outlook for next month.',
    category: 'reports',
    content: {
      app: 'docs',
      html: [
        '<h1>Monthly Sales Report: [Month Year]</h1>',
        '<p>Prepared by: [Name] · Region or team: [Region or Team]</p>',
        '<h2>Summary</h2>',
        '<p>Total revenue: [Amount] · Target: [Amount] · Change vs. last month: [Percent or Amount]</p>',
        '<h2>Sales by Product</h2>',
        '<table><tr><th>Product</th><th>Units</th><th>Revenue</th></tr><tr><td>[Product 1]</td><td>[Units]</td><td>[Amount]</td></tr><tr><td>[Product 2]</td><td>[Units]</td><td>[Amount]</td></tr></table>',
        '<h2>Sales by Channel</h2>',
        '<ul><li>[Channel 1]: [Revenue and share]</li><li>[Channel 2]: [Revenue and share]</li></ul>',
        '<h2>Highlights</h2>',
        '<ul><li>[Win, new customer, or standout result]</li><li>[Win, new customer, or standout result]</li></ul>',
        '<h2>Challenges</h2>',
        '<p>[What held sales back, and what is being done about it.]</p>',
        '<h2>Next Month Outlook</h2>',
        '<p>[Forecast and the actions planned to reach it.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'report-annual-summary',
    title: 'Annual Summary Report',
    description:
      'Year-end summary report covering achievements, financial highlights, lessons learned, and goals for next year.',
    category: 'reports',
    content: {
      app: 'docs',
      html: [
        '<h1>Annual Summary Report: [Year]</h1>',
        '<p>Organization: [Organization Name] · Prepared by: [Name]</p>',
        '<h2>Year in Review</h2>',
        '<p>[Two or three sentences on the year as a whole: what changed and what mattered most.]</p>',
        '<h2>Key Achievements</h2>',
        '<ul><li>[Achievement 1]</li><li>[Achievement 2]</li><li>[Achievement 3]</li></ul>',
        '<h2>Financial Highlights</h2>',
        '<ul><li>Revenue: [Amount]</li><li>Expenses: [Amount]</li><li>Growth vs. prior year: [Percent]</li></ul>',
        '<h2>Challenges and Lessons</h2>',
        '<p>[What was difficult, what was learned, and what will change.]</p>',
        '<h2>Looking Ahead</h2>',
        '<ul><li>[Goal 1 for next year]</li><li>[Goal 2 for next year]</li><li>[Goal 3 for next year]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'report-research',
    title: 'Research Report',
    description:
      'Structured research report with a question, method, findings, analysis, and conclusion, plus a sources section.',
    category: 'reports',
    content: {
      app: 'docs',
      html: [
        '<h1>Research Report: [Topic]</h1>',
        '<p>Author: [Name] · Date: [Date]</p>',
        '<h2>Research Question</h2>',
        '<p>[The question this report sets out to answer.]</p>',
        '<h2>Method</h2>',
        '<p>[How the research was done: sources consulted, data gathered, and approach used.]</p>',
        '<h2>Findings</h2>',
        '<h3>[Finding 1]</h3>',
        '<p>[Evidence and detail supporting this finding.]</p>',
        '<h3>[Finding 2]</h3>',
        '<p>[Evidence and detail supporting this finding.]</p>',
        '<h2>Analysis</h2>',
        '<p>[What the findings mean together, and where they agree or conflict.]</p>',
        '<h2>Conclusion</h2>',
        '<p>[The answer to the research question, and what should happen next.]</p>',
        '<h2>Sources</h2>',
        '<ul><li>[Source 1]</li><li>[Source 2]</li><li>[Source 3]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'budget-monthly-personal',
    title: 'Monthly Personal Budget',
    description:
      'Monthly household budget comparing budgeted and actual spending by category, with the difference on every line.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Monthly Budget',
          headerRow: true,
          widths: [42, 14, 14, 16],
          rows: [
            ['Category', 'Budgeted', 'Actual', 'Difference'],
            ['Housing', 1200, 1200, { formula: 'B2-C2', cached: 0, format: 'currency' }],
            ['Groceries', 400, 385, { formula: 'B3-C3', cached: 15, format: 'currency' }],
            ['Transport', 200, 210, { formula: 'B4-C4', cached: -10, format: 'currency' }],
            ['Utilities', 150, 145, { formula: 'B5-C5', cached: 5, format: 'currency' }],
            ['Entertainment', 100, 95, { formula: 'B6-C6', cached: 5, format: 'currency' }],
            [
              'Total',
              { formula: 'SUM(B2:B6)', cached: 2050, format: 'currency', bold: true },
              { formula: 'SUM(C2:C6)', cached: 2035, format: 'currency', bold: true },
              { formula: 'B7-C7', cached: 15, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'budget-annual-business',
    title: 'Annual Business Budget',
    description:
      'Annual business budget by quarter, with per-category annual totals and a grand total across all quarters.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Annual Budget',
          headerRow: true,
          widths: [24, 14, 14, 14, 14, 16],
          rows: [
            ['Category', 'Q1', 'Q2', 'Q3', 'Q4', 'Annual Total'],
            [
              'Rent',
              6000,
              6000,
              6000,
              6000,
              { formula: 'SUM(B2:E2)', cached: 24000, format: 'currency' },
            ],
            [
              'Payroll',
              15000,
              15000,
              15000,
              15000,
              { formula: 'SUM(B3:E3)', cached: 60000, format: 'currency' },
            ],
            [
              'Marketing',
              2000,
              2000,
              2000,
              2000,
              { formula: 'SUM(B4:E4)', cached: 8000, format: 'currency' },
            ],
            [
              'Supplies',
              1000,
              1200,
              900,
              1100,
              { formula: 'SUM(B5:E5)', cached: 4200, format: 'currency' },
            ],
            [
              'Total',
              { formula: 'SUM(B2:B5)', cached: 24000, format: 'currency', bold: true },
              { formula: 'SUM(C2:C5)', cached: 24200, format: 'currency', bold: true },
              { formula: 'SUM(D2:D5)', cached: 23900, format: 'currency', bold: true },
              { formula: 'SUM(E2:E5)', cached: 24100, format: 'currency', bold: true },
              { formula: 'SUM(F2:F5)', cached: 96200, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'budget-event',
    title: 'Event Budget',
    description:
      'Event budget tracking estimated versus actual costs per item, so overruns show up line by line.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Event Budget',
          headerRow: true,
          widths: [42, 14, 14, 16],
          rows: [
            ['Item', 'Estimated', 'Actual', 'Difference'],
            ['Venue', 2000, 2000, { formula: 'B2-C2', cached: 0, format: 'currency' }],
            ['Catering', 1500, 1620, { formula: 'B3-C3', cached: -120, format: 'currency' }],
            ['Decorations', 400, 350, { formula: 'B4-C4', cached: 50, format: 'currency' }],
            ['Entertainment', 800, 800, { formula: 'B5-C5', cached: 0, format: 'currency' }],
            [
              'Total',
              { formula: 'SUM(B2:B5)', cached: 4700, format: 'currency', bold: true },
              { formula: 'SUM(C2:C5)', cached: 4770, format: 'currency', bold: true },
              { formula: 'B6-C6', cached: -70, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'budget-travel',
    title: 'Travel Budget',
    description:
      'Trip budget comparing planned and actual spending on flights, lodging, meals, and local transport.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Travel Budget',
          headerRow: true,
          widths: [42, 14, 14, 16],
          rows: [
            ['Expense', 'Budgeted', 'Actual', 'Difference'],
            ['Flights', 600, 620, { formula: 'B2-C2', cached: -20, format: 'currency' }],
            ['Hotel', 900, 850, { formula: 'B3-C3', cached: 50, format: 'currency' }],
            ['Meals', 300, 340, { formula: 'B4-C4', cached: -40, format: 'currency' }],
            ['Local transport', 150, 120, { formula: 'B5-C5', cached: 30, format: 'currency' }],
            [
              'Total',
              { formula: 'SUM(B2:B5)', cached: 1950, format: 'currency', bold: true },
              { formula: 'SUM(C2:C5)', cached: 1930, format: 'currency', bold: true },
              { formula: 'B6-C6', cached: 20, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'expense-tracker',
    title: 'Expense Tracker',
    description:
      'Simple running expense log with date, category, description, and amount, plus a total of all entries.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Expenses',
          headerRow: true,
          widths: [16, 24, 42, 14],
          rows: [
            ['Date', 'Category', 'Description', 'Amount'],
            ['[Date]', 'Groceries', '[Store or purchase description]', 45],
            ['[Date]', 'Transport', '[Fuel or fare description]', 12.5],
            ['[Date]', 'Dining', '[Restaurant or meal description]', 85],
            ['[Date]', 'Household', '[Item or service description]', 30],
            [
              '',
              '',
              'Total',
              { formula: 'SUM(D2:D5)', cached: 172.5, format: 'currency', bold: true },
            ],
          ],
        },
      ],
    },
  },
  {
    id: 'schedule-weekly-planner',
    title: 'Weekly Planner',
    description:
      'Week-at-a-glance planner with one focus, top task, and notes for each day of the working week.',
    category: 'schedules',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Weekly Planner',
          headerRow: true,
          widths: [16, 28, 42, 28],
          rows: [
            ['Day', 'Focus', 'Top Task', 'Notes'],
            ['Monday', '[Focus area]', '[Most important task]', '[Notes]'],
            ['Tuesday', '[Focus area]', '[Most important task]', '[Notes]'],
            ['Wednesday', '[Focus area]', '[Most important task]', '[Notes]'],
            ['Thursday', '[Focus area]', '[Most important task]', '[Notes]'],
            ['Friday', '[Focus area]', '[Most important task]', '[Notes]'],
          ],
        },
      ],
    },
  },
  {
    id: 'schedule-project-timeline',
    title: 'Project Timeline',
    description:
      'Project timeline listing each phase with start and end dates, the owner, and the current status.',
    category: 'schedules',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Timeline',
          headerRow: true,
          widths: [28, 16, 16, 20, 16],
          rows: [
            ['Phase', 'Start', 'End', 'Owner', 'Status'],
            ['Discovery', '[Start date]', '[End date]', '[Name]', 'Not started'],
            ['Planning', '[Start date]', '[End date]', '[Name]', 'Not started'],
            ['Build', '[Start date]', '[End date]', '[Name]', 'Not started'],
            ['Testing', '[Start date]', '[End date]', '[Name]', 'Not started'],
            ['Launch', '[Start date]', '[End date]', '[Name]', 'Not started'],
          ],
        },
      ],
    },
  },
  {
    id: 'schedule-shift-rota',
    title: 'Employee Shift Rota',
    description:
      'Weekly shift rota showing who works each day, their role, and their shift times for the whole week.',
    category: 'schedules',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Shift Rota',
          headerRow: true,
          widths: [24, 20, 16, 16, 16, 16, 16],
          rows: [
            ['Employee', 'Role', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
            ['[Employee 1]', '[Role]', '9am-5pm', '9am-5pm', 'Off', '9am-5pm', '9am-5pm'],
            ['[Employee 2]', '[Role]', 'Off', '9am-5pm', '9am-5pm', '9am-5pm', 'Off'],
            ['[Employee 3]', '[Role]', '1pm-9pm', 'Off', '1pm-9pm', 'Off', '1pm-9pm'],
          ],
        },
      ],
    },
  },
  {
    id: 'schedule-content-calendar',
    title: 'Content Calendar',
    description:
      'Content publishing calendar with dates, channels, topics, and status so a whole month of posts stays visible.',
    category: 'schedules',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Content Calendar',
          headerRow: true,
          widths: [16, 20, 42, 16],
          rows: [
            ['Date', 'Channel', 'Topic', 'Status'],
            ['[Date]', '[Blog / Social / Email]', '[Post topic or title]', 'Planned'],
            ['[Date]', '[Blog / Social / Email]', '[Post topic or title]', 'Draft'],
            ['[Date]', '[Blog / Social / Email]', '[Post topic or title]', 'Planned'],
            ['[Date]', '[Blog / Social / Email]', '[Post topic or title]', 'Published'],
          ],
        },
      ],
    },
  },
  {
    id: 'schedule-study-timetable',
    title: 'Study Timetable',
    description:
      'Weekly study timetable organizing subjects, times, and topics to cover for each study session.',
    category: 'schedules',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Study Timetable',
          headerRow: true,
          widths: [16, 24, 18, 42],
          rows: [
            ['Day', 'Subject', 'Time', 'Topic'],
            ['Monday', '[Subject]', '[Start time] – [End time]', '[Topic to study]'],
            ['Tuesday', '[Subject]', '[Start time] – [End time]', '[Topic to study]'],
            ['Wednesday', '[Subject]', '[Start time] – [End time]', '[Topic to study]'],
            ['Thursday', '[Subject]', '[Start time] – [End time]', '[Topic to study]'],
            ['Friday', '[Subject]', '[Start time] – [End time]', '[Topic to study]'],
          ],
        },
      ],
    },
  },
  {
    id: 'presentation-sales',
    title: 'Sales Presentation',
    description:
      'Customer-facing sales deck: the problem, the offer, proof it works, pricing, and a clear next step.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Product or Service Name]',
          subtitle: '[One-line promise: the outcome your customer gets]',
          bullets: [],
        },
        {
          title: 'The Problem',
          bullets: [
            '[The pain your customer feels today.]',
            '[What that pain costs them in time, money, or missed opportunity.]',
            '[Why solving it now matters.]',
          ],
        },
        {
          title: 'Our Solution',
          bullets: [
            '[What you offer, in one sentence.]',
            '[The three features customers value most.]',
            '[How it fits into their existing workflow.]',
          ],
        },
        {
          title: 'Proof It Works',
          bullets: [
            '[Customer result with a number.]',
            '[Testimonial-style quote from a customer.]',
            '[A second result or case study.]',
          ],
        },
        {
          title: 'Pricing and Next Steps',
          bullets: [
            '[Your plans and what each includes.]',
            '[The plan most customers start with, and why.]',
            '[The next step: trial, demo, or call to book.]',
          ],
        },
      ],
    },
  },
  {
    id: 'presentation-training',
    title: 'Training Presentation',
    description:
      'Training deck with objectives, core concepts, a worked example, practice, and a recap to close the session.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Training Topic]',
          subtitle: '[Who this training is for and what they will be able to do afterwards]',
          bullets: [],
        },
        {
          title: 'What You Will Learn',
          bullets: [
            '[Learning objective 1.]',
            '[Learning objective 2.]',
            '[Learning objective 3.]',
          ],
        },
        {
          title: 'Key Concepts',
          bullets: [
            '[Concept 1, explained in one line.]',
            '[Concept 2, explained in one line.]',
            '[Concept 3, explained in one line.]',
          ],
        },
        {
          title: 'Worked Example',
          bullets: [
            '[The scenario or problem to work through.]',
            '[Step one of the solution.]',
            '[Step two, and the result.]',
          ],
        },
        {
          title: 'Practice and Recap',
          bullets: [
            '[Exercise for learners to try now.]',
            '[The three takeaways to remember.]',
            '[Where to find help or follow-up material.]',
          ],
        },
      ],
    },
  },
  {
    id: 'presentation-project-status',
    title: 'Project Status Deck',
    description:
      'Stakeholder status deck: where the project stands, what shipped, what is at risk, and what comes next.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Project Name] Status Update',
          subtitle: '[Reporting period] · [Overall status: On Track, At Risk, or Off Track]',
          bullets: [],
        },
        {
          title: 'Where We Stand',
          bullets: [
            '[Overall progress against the plan.]',
            '[Milestones hit since the last update.]',
            '[Milestones still ahead.]',
          ],
        },
        {
          title: 'What Shipped',
          bullets: ['[Delivered item 1 and its impact.]', '[Delivered item 2 and its impact.]'],
        },
        {
          title: 'Risks and Blockers',
          bullets: [
            '[Risk or blocker 1 and the plan to handle it.]',
            '[Risk or blocker 2 and the plan to handle it.]',
            '[Decisions needed from stakeholders.]',
          ],
        },
        {
          title: 'Next Steps',
          bullets: [
            '[Next milestone and its date.]',
            '[What the team is focused on next.]',
            '[What you need from this audience.]',
          ],
        },
      ],
    },
  },
  {
    id: 'presentation-company-overview',
    title: 'Company Overview Deck',
    description:
      'Company introduction deck: who you are, what you do, who you serve, and why customers choose you.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Company Name]',
          subtitle: '[Tagline: what you do and who you do it for]',
          bullets: [],
        },
        {
          title: 'Who We Are',
          bullets: [
            '[When and where the company started.]',
            '[Your mission in one sentence.]',
            '[Team size, locations, or scale.]',
          ],
        },
        {
          title: 'What We Do',
          bullets: [
            '[Product or service line 1.]',
            '[Product or service line 2.]',
            '[Product or service line 3.]',
          ],
        },
        {
          title: 'Who We Serve',
          bullets: [
            '[Your main customer group.]',
            '[A second customer group.]',
            '[What your customers have in common.]',
          ],
        },
        {
          title: 'Why Choose Us',
          bullets: [
            '[Differentiator 1, with proof if you have it.]',
            '[Differentiator 2, with proof if you have it.]',
            '[How to get started with us.]',
          ],
        },
      ],
    },
  },
  {
    id: 'flyer-event',
    title: 'Event Flyer',
    description:
      'Eye-catching event flyer layout with the event name, date, venue, highlights, and how to attend or register.',
    category: 'flyers',
    content: {
      app: 'docs',
      html: [
        '<h1>[Event Name]</h1>',
        '<p><strong>[Date] · [Start Time] – [End Time]</strong><br>[Venue Name], [Venue Address]</p>',
        '<h2>Join Us For</h2>',
        '<p>[One or two sentences describing the event and why people will enjoy it.]</p>',
        '<h2>Highlights</h2>',
        '<ul><li>[Highlight 1]</li><li>[Highlight 2]</li><li>[Highlight 3]</li></ul>',
        '<h2>Admission</h2>',
        '<p>[Free / Ticket price / Registration details]</p>',
        '<h2>Register</h2>',
        '<p>[Registration link, phone, or email] · Questions? Contact [Contact Name] at [Contact Details].</p>',
      ].join('\n'),
    },
  },
  {
    id: 'flyer-business-promo',
    title: 'Business Promo Flyer',
    description:
      'Promotional flyer for a business offer or opening, with the offer, the details, and a call to action.',
    category: 'flyers',
    content: {
      app: 'docs',
      html: [
        '<h1>[Business Name]</h1>',
        '<h2>[Offer Headline: e.g. Grand Opening Special]</h2>',
        '<p>[One or two sentences about your business and what makes it worth visiting.]</p>',
        '<h2>The Offer</h2>',
        '<p><strong>[Offer details: discount, bundle, or special]</strong><br>Valid: [Start Date] – [End Date]</p>',
        '<h2>What We Offer</h2>',
        '<ul><li>[Product or service 1]</li><li>[Product or service 2]</li><li>[Product or service 3]</li></ul>',
        '<h2>Visit Us</h2>',
        '<p>[Address]<br>[Hours]<br>[Phone] · [Website]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'flyer-community-newsletter',
    title: 'Community Newsletter',
    description:
      'Single-page community newsletter with announcements, upcoming events, and a spotlight on a local highlight.',
    category: 'flyers',
    content: {
      app: 'docs',
      html: [
        '<h1>[Community Name] Newsletter</h1>',
        '<p>[Month Year] · [Issue or edition, if used]</p>',
        '<h2>Announcements</h2>',
        '<ul><li>[Announcement 1]</li><li>[Announcement 2]</li></ul>',
        '<h2>Upcoming Events</h2>',
        '<table><tr><th>Date</th><th>Event</th><th>Location</th></tr><tr><td>[Date]</td><td>[Event 1]</td><td>[Location]</td></tr><tr><td>[Date]</td><td>[Event 2]</td><td>[Location]</td></tr></table>',
        '<h2>Community Spotlight</h2>',
        '<p>[A person, group, or project worth celebrating, and why.]</p>',
        '<h2>Get Involved</h2>',
        '<p>[How readers can volunteer, attend, or contribute.] Contact [Contact Name] at [Contact Details].</p>',
      ].join('\n'),
    },
  },
  {
    id: 'form-job-application',
    title: 'Job Application Form',
    description:
      'Printable job application form covering personal details, position sought, work history, and references.',
    category: 'forms',
    content: {
      app: 'docs',
      html: [
        '<h1>Job Application Form</h1>',
        '<p>Position applying for: ______________________ Date: __________</p>',
        '<h2>Personal Information</h2>',
        '<p>Full name: ______________________<br>Address: ______________________<br>Phone: ______________________ Email: ______________________</p>',
        '<h2>Position Details</h2>',
        '<p>Desired role: ______________________ Available start date: __________<br>Work authorization: [Yes / No] · Willing to relocate: [Yes / No]</p>',
        '<h2>Employment History</h2>',
        '<table><tr><th>Employer</th><th>Role</th><th>Dates</th><th>Reason for Leaving</th></tr><tr><td>______________________</td><td>______________________</td><td>__________</td><td>______________________</td></tr><tr><td>______________________</td><td>______________________</td><td>__________</td><td>______________________</td></tr></table>',
        '<h2>Education</h2>',
        '<p>School: ______________________ Qualification: ______________________ Year: __________</p>',
        '<h2>References</h2>',
        '<p>Name: ______________________ Relationship: ______________________ Phone: ______________________<br>Name: ______________________ Relationship: ______________________ Phone: ______________________</p>',
        '<p>Signature: ______________________ Date: __________</p>',
      ].join('\n'),
    },
  },
  {
    id: 'form-feedback-survey',
    title: 'Feedback Survey Form',
    description:
      'Customer or attendee feedback survey with rating questions, open questions, and an optional contact section.',
    category: 'forms',
    content: {
      app: 'docs',
      html: [
        '<h1>Feedback Survey: [Product, Service, or Event Name]</h1>',
        '<p>Thank you for sharing your feedback. Your answers help us improve.</p>',
        '<h2>Overall Experience</h2>',
        '<p>How would you rate your overall experience? [1 – Poor, 2 – Fair, 3 – Good, 4 – Very Good, 5 – Excellent]</p>',
        '<p>How likely are you to recommend us to others? [1 – Not at all, to 5 – Very likely]</p>',
        '<h2>What Worked Well</h2>',
        '<p>________________________________________________________________________</p>',
        '<h2>What Could Be Better</h2>',
        '<p>________________________________________________________________________</p>',
        '<h2>Additional Comments</h2>',
        '<p>________________________________________________________________________</p>',
        '<h2>Contact (Optional)</h2>',
        '<p>Name: ______________________ Email: ______________________<br>[Tick here if you are happy for us to follow up on your feedback.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'form-event-planning-checklist',
    title: 'Event Planning Checklist',
    description:
      'Step-by-step event planning checklist from venue and vendors through promotion, setup, and follow-up.',
    category: 'forms',
    content: {
      app: 'docs',
      html: [
        '<h1>Event Planning Checklist: [Event Name]</h1>',
        '<p>Event date: [Date] · Venue: [Venue] · Lead planner: [Name]</p>',
        '<h2>Venue and Logistics</h2>',
        '<ul><li>[ ] Book venue and confirm capacity and layout</li><li>[ ] Arrange seating, tables, and signage</li><li>[ ] Confirm parking and accessibility</li></ul>',
        '<h2>Vendors</h2>',
        '<ul><li>[ ] Book caterer and confirm menu</li><li>[ ] Book photographer or videographer</li><li>[ ] Book entertainment or speakers</li></ul>',
        '<h2>Promotion</h2>',
        '<ul><li>[ ] Send invitations or publish registration</li><li>[ ] Post on social media and community boards</li><li>[ ] Send a reminder one week before</li></ul>',
        '<h2>On the Day</h2>',
        '<ul><li>[ ] Arrive early for setup and sound check</li><li>[ ] Brief volunteers and vendors</li><li>[ ] Prepare a welcome and registration table</li></ul>',
        '<h2>After the Event</h2>',
        '<ul><li>[ ] Send thank-you messages</li><li>[ ] Collect feedback</li><li>[ ] Settle final invoices</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'lesson-plan',
    title: 'Lesson Plan',
    description:
      'Classroom lesson plan with objectives, materials, timed activities, and an assessment to check understanding.',
    category: 'education',
    content: {
      app: 'docs',
      html: [
        '<h1>Lesson Plan: [Lesson Title]</h1>',
        '<p>Subject: [Subject] · Grade or level: [Level] · Duration: [Minutes] · Date: [Date]</p>',
        '<h2>Learning Objectives</h2>',
        '<p>By the end of this lesson, students will be able to:</p>',
        '<ul><li>[Objective 1]</li><li>[Objective 2]</li><li>[Objective 3]</li></ul>',
        '<h2>Materials</h2>',
        '<ul><li>[Material 1]</li><li>[Material 2]</li></ul>',
        '<h2>Lesson Activities</h2>',
        '<table><tr><th>Time</th><th>Activity</th><th>Notes</th></tr><tr><td>[Minutes]</td><td>Introduction: [How you open the lesson]</td><td>[Notes]</td></tr><tr><td>[Minutes]</td><td>Main activity: [What students do]</td><td>[Notes]</td></tr><tr><td>[Minutes]</td><td>Wrap-up: [How you close the lesson]</td><td>[Notes]</td></tr></table>',
        '<h2>Assessment</h2>',
        '<p>[How you will check understanding: questions, quiz, exit ticket, or task.]</p>',
        '<h2>Differentiation</h2>',
        '<p>[Support for students who need it, and extension work for students who finish early.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'course-syllabus',
    title: 'Course Syllabus',
    description:
      'Course syllabus with an overview, schedule, grading policy, and expectations for students and instructors.',
    category: 'education',
    content: {
      app: 'docs',
      html: [
        '<h1>Course Syllabus: [Course Title]</h1>',
        '<p>Instructor: [Name] · Term: [Term and Year] · Meeting times: [Days and Times] · Location: [Room or Online Link]</p>',
        '<h2>Course Overview</h2>',
        '<p>[What the course covers, and what students will gain from it.]</p>',
        '<h2>Learning Outcomes</h2>',
        '<ul><li>[Outcome 1]</li><li>[Outcome 2]</li><li>[Outcome 3]</li></ul>',
        '<h2>Schedule</h2>',
        '<table><tr><th>Week</th><th>Topic</th><th>Reading or Assignment</th></tr><tr><td>1</td><td>[Topic]</td><td>[Reading or assignment]</td></tr><tr><td>2</td><td>[Topic]</td><td>[Reading or assignment]</td></tr></table>',
        '<h2>Grading</h2>',
        '<table><tr><th>Component</th><th>Weight</th></tr><tr><td>[Assignments]</td><td>[Percent]</td></tr><tr><td>[Exams]</td><td>[Percent]</td></tr><tr><td>[Participation]</td><td>[Percent]</td></tr></table>',
        '<h2>Policies</h2>',
        '<p>[Attendance, late work, academic integrity, and communication expectations.]</p>',
      ].join('\n'),
    },
  },
  {
    id: 'essay-assignment-structure',
    title: 'Essay and Report Assignment Structure',
    description:
      'Assignment outline template structuring an essay or report: thesis, body sections, evidence, and conclusion.',
    category: 'education',
    content: {
      app: 'docs',
      html: [
        '<h1>Essay and Report Assignment: [Assignment Title]</h1>',
        '<p>Subject: [Subject] · Due date: [Date] · Length: [Word count or pages]</p>',
        '<h2>Assignment Question</h2>',
        '<p>[The question or prompt the essay or report must answer.]</p>',
        '<h2>Thesis Statement</h2>',
        '<p>[Your central argument in one or two sentences.]</p>',
        '<h2>Introduction</h2>',
        '<p>[Background, why the topic matters, and a preview of your argument.]</p>',
        '<h2>Body Section 1: [Heading]</h2>',
        '<p>[Main point, evidence, and analysis.]</p>',
        '<h2>Body Section 2: [Heading]</h2>',
        '<p>[Main point, evidence, and analysis.]</p>',
        '<h2>Counterargument (if required)</h2>',
        '<p>[An opposing view, and your response to it.]</p>',
        '<h2>Conclusion</h2>',
        '<p>[Restate the argument, summarize the evidence, and state the wider significance.]</p>',
        '<h2>References</h2>',
        '<ul><li>[Source 1]</li><li>[Source 2]</li></ul>',
      ].join('\n'),
    },
  },
  {
    id: 'gradebook-tracker',
    title: 'Gradebook Tracker',
    description:
      'Gradebook tracking scores across three assignments per student, with each student average calculated automatically.',
    category: 'education',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Gradebook',
          headerRow: true,
          widths: [28, 16, 16, 16, 16],
          rows: [
            ['Student', 'Assignment 1', 'Assignment 2', 'Assignment 3', 'Average'],
            ['[Student 1]', 90, 84, 78, { formula: 'AVERAGE(B2:D2)', cached: 84, format: 'int' }],
            ['[Student 2]', 75, 90, 90, { formula: 'AVERAGE(B3:D3)', cached: 85, format: 'int' }],
            ['[Student 3]', 88, 92, 84, { formula: 'AVERAGE(B4:D4)', cached: 88, format: 'int' }],
          ],
        },
      ],
    },
  },
  // Team-format templates adapted from the structural shapes of a colleague's
  // onboarding pack (working-preferences guide, multi-page deck layouts,
  // one-page update slide, navy table formats). All wording and figures here
  // are original placeholders — no names, employers, real numbers, or brand
  // assets from the source were carried over.
  {
    id: 'team-working-preferences',
    title: 'Team Working Preferences',
    description:
      'One-page guide to how a team prefers to communicate and run projects: short complete emails, bullets for multi-point messages, the right channel for each kind of message, and a consistent project workflow.',
    category: 'meetings',
    content: {
      app: 'docs',
      html: [
        '<h1>Team Working Preferences</h1>',
        '<p>[This guide sets out how our team prefers to work together. Replace the examples with your own team norms.]</p>',
        '<h2>Email</h2>',
        '<ul><li>Keep emails short, but include everything the reader needs to act — no missing context.</li><li>When forwarding a message, add a note explaining what is being forwarded and what you need from the reader. Never send a forward with no explanation.</li><li>When a message has several points or questions, list them as bullets so each one can be answered separately. Bulleted messages get faster replies than long paragraphs.</li><li>Long paragraph emails tend to be set aside for later, so break them up.</li></ul>',
        '<h2>Choosing the channel</h2>',
        '<ul><li>Use email for deliverables, questions, and requests, so they can be tracked and answered from a desk.</li><li>Use text messages for scheduling, quick or urgent items, and casual conversation.</li><li>When a topic needs back-and-forth discussion, use a short call of about 15 minutes instead of a long email thread.</li></ul>',
        '<h2>Project working style</h2>',
        '<ul><li>Start by looking at other examples for ideas, sketch the design on paper, then begin building. The sketch step leads to better results.</li><li>Agree an estimated deadline at the start of each project so everyone plans around the same target.</li><li>Close each project with a final review call to sign off the work, so nothing lingers unfinished.</li></ul>',
        '<h2>Deliverable formats</h2>',
        '<p>Keep deliverables consistent: tables go in the spreadsheet format, short updates go in the one-page format, and longer presentations go in the multi-page format.</p>',
      ].join('\n'),
    },
  },
  {
    id: 'executive-presentation',
    title: 'Executive Presentation',
    description:
      'Six-slide executive deck distilled from multi-page presentation layouts: title, agenda, section divider, headline statement, summary, and next steps.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Presentation Title]',
          subtitle: '[A one-line statement of the topic and the decision or outcome you want]',
          bullets: [],
        },
        {
          title: 'Agenda',
          bullets: [
            '[Agenda item 1]',
            '[Agenda item 2]',
            '[Agenda item 3]',
            '[Agenda item 4]',
            '[Agenda item 5]',
          ],
        },
        {
          title: 'Section 1: [Section Title]',
          bullets: ['[What this section covers, in one sentence.]'],
        },
        {
          title: '[Headline Statement]',
          bullets: [
            '[Supporting point for the headline.]',
            '[Supporting point for the headline.]',
            '[Supporting point for the headline.]',
          ],
        },
        {
          title: 'Summary',
          bullets: [
            '[Recap of the first key point.]',
            '[Recap of the second key point.]',
            '[One-sentence summary statement.]',
          ],
        },
        {
          title: 'Next Steps',
          bullets: [
            '[Next step, owner, and timing.]',
            '[Next step, owner, and timing.]',
            '[What you need from the audience.]',
          ],
        },
      ],
    },
  },
  {
    id: 'one-page-project-update',
    title: 'One-Page Project Update',
    description:
      'Project update in the one-page shape: header, highlight sentences, key points, and a summary statement on the first slide, with backup detail slides behind it.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Project Name] — Project Update',
          subtitle: '[Reporting period]',
          bullets: [
            '[Highlight sentence: the most important progress this period.]',
            '[Highlight sentence: the second most important progress this period.]',
            'Key point: [Point and supporting sentence.]',
            'Key point: [Point and supporting sentence.]',
            'Key point: [Point and supporting sentence.]',
            'Summary: [One-sentence summary statement.]',
            'Headline figure: [Figure, e.g. budget used or milestones reached]',
          ],
        },
        {
          title: 'What Changed Since Last Update',
          bullets: [
            '[Change since the last update, and why it matters.]',
            '[Change since the last update, and why it matters.]',
            '[Change since the last update, and why it matters.]',
          ],
        },
        {
          title: 'Risks and Next Steps',
          bullets: [
            '[Risk or open issue, and the plan to address it.]',
            '[Risk or open issue, and the plan to address it.]',
            '[Next step, owner, and timing.]',
          ],
        },
      ],
    },
  },
  {
    id: 'financial-review',
    title: 'Financial Review',
    description:
      'Financial review deck: results against plan, what drove the variance, and the forward outlook — all figures are placeholders for your own numbers.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Period] Financial Review',
          subtitle: '[One-line summary: ahead of, behind, or in line with plan]',
          bullets: [],
        },
        {
          title: 'Results Against Plan',
          bullets: [
            'Planned: [Planned amount]',
            'Actual: [Actual amount]',
            'Variance: [Variance amount and percent]',
            '[One sentence on what the variance means.]',
          ],
        },
        {
          title: 'What Drove the Variance',
          bullets: [
            '[Driver 1 and its effect on the result.]',
            '[Driver 2 and its effect on the result.]',
            '[Driver 3 and its effect on the result.]',
          ],
        },
        {
          title: 'Outlook',
          bullets: [
            '[Expected result for the next period.]',
            '[Assumption the outlook depends on.]',
            '[Action being taken to close any gap.]',
          ],
        },
      ],
    },
  },
  {
    id: 'process-timeline',
    title: 'Process Timeline',
    description:
      'Generic process timeline: near-term, mid-term, and long-term phases, milestone rows, and a weekly execution plan.',
    category: 'presentations',
    content: {
      app: 'slides',
      slides: [
        {
          title: '[Process Name] Timeline',
          subtitle: '[What this process delivers, and by when]',
          bullets: [],
        },
        {
          title: 'Phases',
          bullets: [
            'Near-term: [What happens in this phase.]',
            'Mid-term: [What happens in this phase.]',
            'Long-term: [What happens in this phase.]',
          ],
        },
        {
          title: 'Milestones',
          bullets: [
            '[Milestone 1] — [Target timing]',
            '[Milestone 2] — [Target timing]',
            '[Milestone 3] — [Target timing]',
            '[Milestone 4] — [Target timing]',
          ],
        },
        {
          title: 'Weekly Execution Plan',
          bullets: [
            '[Week 1]: [Tasks for this week.]',
            '[Week 2]: [Tasks for this week.]',
            '[Week 3]: [Tasks for this week.]',
            '[Week 4]: [Tasks for this week.]',
          ],
        },
      ],
    },
  },
  {
    id: 'formatted-data-table',
    title: 'Formatted Data Table',
    description:
      'Styled data table with a bold header row, sample rows, and an automatic total, plus a comparison sheet for weighing options side by side.',
    category: 'budgets',
    content: {
      app: 'sheets',
      sheets: [
        {
          name: 'Data Table',
          headerRow: true,
          widths: [28, 44, 16, 28],
          rows: [
            ['Item', 'Description', 'Amount', 'Notes'],
            ['[Item 1]', '[Short description of the item]', 120, '[Note]'],
            ['[Item 2]', '[Short description of the item]', 250, '[Note]'],
            ['[Item 3]', '[Short description of the item]', 80, '[Note]'],
            [
              'Total',
              '',
              { formula: 'SUM(C2:C4)', cached: 450, format: 'currency', bold: true },
              '',
            ],
          ],
        },
        {
          name: 'Comparison',
          headerRow: true,
          widths: [28, 16, 16, 32],
          rows: [
            ['Option', 'Cost', 'Rating', 'Notes'],
            ['[Option 1]', 100, 'High', '[Note]'],
            ['[Option 2]', 150, 'Medium', '[Note]'],
            ['[Option 3]', 75, 'Low', '[Note]'],
          ],
        },
      ],
    },
  },
]
