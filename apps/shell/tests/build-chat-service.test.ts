import { mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { BuildChatService, extractBuildJson } from '../src/main/build-chat-service'
import type { CliRunner } from '../src/main/mcp/cli-runner'

/**
 * Build chat service: the clarify gate, the edit-confirmation gate, plan
 * execution against the REAL template builders (Sheets/Slides bytes are
 * produced by template-files into a temp folder), the no-overwrite rule,
 * and the deterministic no-model fallback. The model is a scripted fake;
 * the docx writer is a fake (the real one shells to the bundled CLI) —
 * everything else on the create path is production code.
 */

const cliStub = {
  run: async () => ({ ok: false, json: null }),
} as unknown as CliRunner

interface Harness {
  service: BuildChatService
  saveDir: string
  openedFiles: string[]
  openedDocsTemplates: Array<{ title: string; html: string }>
  writtenDocx: Array<{ markdown: string; out: string }>
}

async function makeHarness(planReplies?: Array<string | null>): Promise<Harness> {
  const userDataPath = await mkdtemp(join(tmpdir(), 'build-chat-user-'))
  const saveDir = await mkdtemp(join(tmpdir(), 'build-chat-save-'))
  const openedFiles: string[] = []
  const openedDocsTemplates: Array<{ title: string; html: string }> = []
  const writtenDocx: Array<{ markdown: string; out: string }> = []
  const queue = planReplies ? [...planReplies] : null
  const service = new BuildChatService({
    userDataPath,
    cli: cliStub,
    saveDir: () => saveDir,
    openFile: (path) => {
      openedFiles.push(path)
    },
    openDocsTemplate: (input) => {
      openedDocsTemplates.push(input)
    },
    ...(queue
      ? {
          planWithAi: async () => (queue.length > 0 ? (queue.shift() ?? null) : null),
        }
      : {}),
    writeDocxFile: async (markdown, out) => {
      writtenDocx.push({ markdown, out })
      await writeFile(out, markdown, 'utf8')
    },
  })
  return { service, saveDir, openedFiles, openedDocsTemplates, writtenDocx }
}

async function writeDocxFixture(path: string, texts: string[]): Promise<void> {
  const zip = new JSZip()
  const paragraphs = texts.map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document><w:body>${paragraphs}</w:body></w:document>`,
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function readDocxXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const entry = zip.file('word/document.xml')
  if (!entry) throw new Error('fixture docx has no document.xml')
  return entry.async('string')
}

const INVOICE_PLAN = JSON.stringify({
  plan: [
    {
      kind: 'create',
      app: 'sheets',
      title: 'Invoice for Acme',
      templateId: 'invoice-service',
    },
  ],
  message: 'Making the invoice.',
})

describe('extractBuildJson', () => {
  it('reads bare and fenced JSON', () => {
    expect(extractBuildJson('{"questions":["a"]}')).toEqual({ questions: ['a'] })
    expect(extractBuildJson('```json\n{"plan":[]}\n```')).toEqual({ plan: [] })
    expect(extractBuildJson('no json here')).toBeNull()
  })
})

describe('the clarify gate', () => {
  it('a questions reply builds nothing; the plan builds after answers', async () => {
    const harness = await makeHarness([
      JSON.stringify({
        questions: ['Who is the invoice for?', 'What are the line items?'],
        message: 'Two details first.',
      }),
      INVOICE_PLAN,
    ])
    const asked = await harness.service.send('I need an invoice')
    expect(asked.kind).toBe('questions')
    expect(asked.questions).toEqual(['Who is the invoice for?', 'What are the line items?'])
    // the gate: no file, no tab, no docx write happened on a questions turn
    expect(await readdir(harness.saveDir)).toEqual([])
    expect(harness.openedFiles).toEqual([])
    expect(harness.writtenDocx).toEqual([])

    const built = await harness.service.send('Acme Corp; design work')
    expect(built.kind).toBe('built')
    expect(built.files).toHaveLength(1)
    expect(built.files[0]?.path).toBeTruthy()
    expect(existsSync(built.files[0]!.path!)).toBe(true)
    expect(harness.openedFiles).toEqual([built.files[0]!.path])
  })

  it('invalid model output becomes a clarifying question, not a build', async () => {
    const harness = await makeHarness(['Sure thing, building it now!'])
    const turnResult = await harness.service.send('make me something')
    expect(turnResult.kind).toBe('questions')
    expect(turnResult.questions[0]).toContain('tell me a bit more')
    expect(await readdir(harness.saveDir)).toEqual([])
  })
})

describe('plan execution across apps', () => {
  it('creates docs, sheets and slides files plus a filled docs template in one plan', async () => {
    const plan = JSON.stringify({
      plan: [
        {
          kind: 'create',
          app: 'docs',
          title: 'Cover Letter',
          markdown: '# Cover Letter\n\nDear team,',
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
                ['Design', 500],
              ],
            },
          ],
        },
        {
          kind: 'create',
          app: 'slides',
          title: 'Pitch',
          slides: [{ title: 'Pitch', bullets: ['Hello'] }],
        },
        {
          kind: 'create',
          app: 'docs',
          title: 'My Resume',
          templateId: 'resume-modern',
          fill: [{ find: '[Your Full Name]', replace: 'Alex Boles' }],
        },
      ],
      message: 'Creating all four.',
    })
    const harness = await makeHarness([plan])
    const turnResult = await harness.service.send('cover letter, quote, pitch deck and my resume')
    expect(turnResult.kind).toBe('built')
    expect(turnResult.errors).toEqual([])
    expect(turnResult.files).toHaveLength(4)
    expect(turnResult.planLines).toHaveLength(4)

    const [letter, quote, pitch, resumeFile] = turnResult.files
    // custom docs went through the docx writer with the plan's markdown
    expect(letter?.path?.endsWith('Cover Letter.docx')).toBe(true)
    expect(harness.writtenDocx[0]?.markdown).toContain('Dear team,')
    // sheets and slides are real files from the real builders (zip magic)
    for (const file of [quote, pitch]) {
      const bytes = await readFile(file!.path!)
      expect(bytes.subarray(0, 2).toString('latin1')).toBe('PK')
    }
    expect(quote?.path?.endsWith('Quote.xlsx')).toBe(true)
    expect(pitch?.path?.endsWith('Pitch.pptx')).toBe(true)
    // the docs template opened as a tab with the fill applied
    expect(resumeFile?.path).toBeUndefined()
    expect(harness.openedDocsTemplates).toHaveLength(1)
    expect(harness.openedDocsTemplates[0]?.html).toContain('Alex Boles')
    expect(harness.openedDocsTemplates[0]?.html).not.toContain('[Your Full Name]')
    // every file-backed result was opened as a tab
    expect(harness.openedFiles).toEqual([letter?.path, quote?.path, pitch?.path])
  })

  it('never overwrites: a clashing title lands on a unique-suffixed path', async () => {
    const harness = await makeHarness([
      JSON.stringify({
        plan: [
          {
            kind: 'create',
            app: 'sheets',
            title: 'Quote',
            sheets: [{ name: 'Quote', rows: [['Item'], ['Design']] }],
          },
        ],
      }),
    ])
    const existing = join(harness.saveDir, 'Quote.xlsx')
    await writeFile(existing, 'SENTINEL', 'utf8')
    const turnResult = await harness.service.send('make a quote')
    expect(turnResult.kind).toBe('built')
    expect(turnResult.files[0]?.path?.endsWith('Quote-2.xlsx')).toBe(true)
    expect(await readFile(existing, 'utf8')).toBe('SENTINEL')
    const created = await readFile(turnResult.files[0]!.path!)
    expect(created.subarray(0, 2).toString('latin1')).toBe('PK')
  })
})

describe('the edit-confirmation gate', () => {
  const editPlan = (path: string, find: string, replace: string): string =>
    JSON.stringify({
      plan: [{ kind: 'edit', app: 'docs', path, find, replace }],
      message: 'Here is the change.',
    })

  it('an edit waits for an explicit yes, and only then runs', async () => {
    const harness = await makeHarness()
    const letterPath = join(harness.saveDir, 'letter.docx')
    await writeDocxFixture(letterPath, ['Dear Sir', 'Body text.'])
    const service = new BuildChatService({
      userDataPath: harness.saveDir,
      cli: cliStub,
      saveDir: () => harness.saveDir,
      openFile: (path) => harness.openedFiles.push(path),
      openDocsTemplate: (input) => harness.openedDocsTemplates.push(input),
      planWithAi: async () => {
        const replies = [
          editPlan(letterPath, 'Dear Sir', 'Dear Ms. Lee'),
          '{"questions":["What instead?"]}',
          editPlan(letterPath, 'Dear Sir', 'Dear Ms. Lee'),
        ]
        return replies.shift() ?? null
      },
      writeDocxFile: async (markdown, out) => {
        await writeFile(out, markdown, 'utf8')
      },
    })

    const proposed = await service.send('Change Dear Sir to Dear Ms. Lee in my letter')
    expect(proposed.kind).toBe('confirm')
    expect(proposed.planLines[0]).toContain('letter.docx')
    expect(await readDocxXml(letterPath)).toContain('Dear Sir')

    // a non-confirmation drops the plan unexecuted
    const dropped = await service.send('actually, what else can you do?')
    expect(dropped.kind).toBe('questions')
    expect(dropped.message).toContain('Dropped that edit plan.')
    expect(await readDocxXml(letterPath)).toContain('Dear Sir')

    // propose again, confirm, and the edit lands
    const reproposed = await service.send('Change Dear Sir to Dear Ms. Lee in my letter')
    expect(reproposed.kind).toBe('confirm')
    const done = await service.send('yes')
    expect(done.kind).toBe('built')
    expect(done.files).toHaveLength(1)
    const xml = await readDocxXml(letterPath)
    expect(xml).toContain('Dear Ms. Lee')
    expect(xml).not.toContain('Dear Sir')
  })

  it('an edit whose text occurs twice changes nothing and says so', async () => {
    const harness = await makeHarness([])
    const letterPath = join(harness.saveDir, 'totals.docx')
    await writeDocxFixture(letterPath, ['Total', 'Total'])
    const service = new BuildChatService({
      userDataPath: harness.saveDir,
      cli: cliStub,
      saveDir: () => harness.saveDir,
      openFile: () => {},
      openDocsTemplate: () => {},
      planWithAi: async () => editPlan(letterPath, 'Total', 'Sum'),
      writeDocxFile: async (markdown, out) => {
        await writeFile(out, markdown, 'utf8')
      },
    })
    const proposed = await service.send('Rename Total to Sum in totals.docx')
    expect(proposed.kind).toBe('confirm')
    const done = await service.send('yes')
    expect(done.kind).toBe('built')
    expect(done.files).toEqual([])
    expect(done.errors[0]).toContain('not found exactly once')
    const xml = await readDocxXml(letterPath)
    expect(xml).toContain('Total')
    expect(xml).not.toContain('Sum')
  })
})

describe('the deterministic no-model fallback', () => {
  it('reports no-model when nothing matches and no AI is configured', async () => {
    const harness = await makeHarness()
    expect(harness.service.modelAvailable()).toBe(false)
    const turnResult = await harness.service.send('blorp zanzibar quixotic')
    expect(turnResult.kind).toBe('no-model')
    expect(turnResult.modelAvailable).toBe(false)
    expect(turnResult.message).toContain('No AI is set up yet')
  })

  it('builds a placeholder-free template immediately', async () => {
    const harness = await makeHarness()
    const turnResult = await harness.service.send('I need a service invoice')
    expect(turnResult.kind).toBe('built')
    expect(turnResult.files).toHaveLength(1)
    expect(turnResult.files[0]?.path?.endsWith('.xlsx')).toBe(true)
    const bytes = await readFile(turnResult.files[0]!.path!)
    expect(bytes.subarray(0, 2).toString('latin1')).toBe('PK')
  })

  it('asks placeholder questions, then fills from the answers', async () => {
    const harness = await makeHarness()
    const asked = await harness.service.send('I need a payment receipt')
    expect(asked.kind).toBe('questions')
    expect(asked.questions).toEqual([
      'What should I put for [Item 1]?',
      'What should I put for [Item 2]?',
    ])
    expect(await readdir(harness.saveDir)).toEqual([])

    const built = await harness.service.send('Keyboard\nMouse')
    expect(built.kind).toBe('built')
    expect(built.files).toHaveLength(1)
    const zip = await JSZip.loadAsync(await readFile(built.files[0]!.path!))
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheetXml).toContain('Keyboard')
    expect(sheetXml).toContain('Mouse')
    expect(sheetXml).not.toContain('[Item 1]')
  })

  it('re-asks when the answers do not cover the placeholders', async () => {
    const harness = await makeHarness()
    await harness.service.send('I need a payment receipt')
    const again = await harness.service.send('Keyboard')
    expect(again.kind).toBe('questions')
    expect(again.questions).toHaveLength(2)
    expect(await readdir(harness.saveDir)).toEqual([])
  })
})
