/**
 * Builders that turn template definitions into real files: a Sheets
 * template becomes a genuine .xlsx (values, formulas with cached results,
 * header styling, column widths), a Slides template becomes a genuine
 * .pptx through pptx-engine. Docs templates need no builder — the shell
 * feeds their HTML straight into the docs tab's AI-content pipeline.
 *
 * The .xlsx writer mirrors the minimal-workbook shape the app already
 * ships (packages/xlsx-gateway csv-import: inline strings, the default
 * theme, a stylesheet the sidecar's formula engine accepts), extended to
 * multiple sheets, formulas, and a small style table. No new dependency:
 * JSZip is already a shell dependency.
 */

import JSZip from 'jszip'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  savePptx,
} from '@genoffice/pptx-engine'
import { DEFAULT_THEME_XML } from '@genoffice/xlsx-gateway/gateway/xlsx-default-theme'
import { validateSheetName } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { isFormulaCell, type TemplateSheet, type TemplateSlide } from '../shared/template-defs'

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
}

function columnLabel(column: number): string {
  let label = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    label = String.fromCharCode(65 + (remaining % 26)) + label
    remaining = Math.floor(remaining / 26)
  }
  return label
}

// Style table: 0 normal, 1 bold, 2 header (white bold on TechPartner green),
// 3 currency, 4 currency bold, 5 percent, 6 integer.
const TEMPLATE_STYLESHEET_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="3">' +
  '<numFmt numFmtId="164" formatCode="#,##0.00"/>' +
  '<numFmt numFmtId="165" formatCode="0.0%"/>' +
  '<numFmt numFmtId="166" formatCode="#,##0"/>' +
  '</numFmts>' +
  '<fonts count="3">' +
  '<font><sz val="11"/><name val="Calibri"/></font>' +
  '<font><sz val="11"/><b/><name val="Calibri"/></font>' +
  '<font><sz val="11"/><b/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
  '</fonts>' +
  '<fills count="3">' +
  '<fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FF1C5B4A"/><bgColor rgb="FF1C5B4A"/></patternFill></fill>' +
  '</fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="7">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '</cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>'

function styleIndex(cell: unknown, isHeader: boolean): number {
  if (isHeader) return 2
  if (isFormulaCell(cell as never)) {
    const formula = cell as { format?: string; bold?: boolean }
    if (formula.format === 'currency') return formula.bold ? 4 : 3
    if (formula.format === 'percent') return 5
    if (formula.format === 'int') return 6
    return formula.bold ? 1 : 0
  }
  return 0
}

function buildWorksheetXml(sheet: TemplateSheet): string {
  const lines: string[] = []
  let maxColumns = 1
  sheet.rows.forEach((row, rowIndex) => {
    const cells: string[] = []
    const isHeader = sheet.headerRow === true && rowIndex === 0
    row.forEach((cell, columnIndex) => {
      if (cell === '') return
      maxColumns = Math.max(maxColumns, columnIndex + 1)
      const reference = `${columnLabel(columnIndex)}${rowIndex + 1}`
      const style = styleIndex(cell, isHeader)
      const styleAttr = style === 0 ? '' : ` s="${style}"`
      if (isFormulaCell(cell)) {
        // string cached results ride in <v> with t="str"; numbers ride bare
        const cachedAttr = typeof cell.cached === 'number' ? '' : ' t="str"'
        const cachedValue =
          typeof cell.cached === 'number' ? String(cell.cached) : escapeXml(cell.cached)
        cells.push(
          `<c r="${reference}"${styleAttr}${cachedAttr}><f>${escapeXml(cell.formula)}</f><v>${cachedValue}</v></c>`,
        )
      } else if (typeof cell === 'number') {
        cells.push(`<c r="${reference}"${styleAttr}><v>${cell}</v></c>`)
      } else {
        cells.push(
          `<c r="${reference}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(cell)}</t></is></c>`,
        )
      }
    })
    if (cells.length > 0) lines.push(`<row r="${rowIndex + 1}">${cells.join('')}</row>`)
  })
  const cols =
    sheet.widths && sheet.widths.length > 0
      ? `<cols>${sheet.widths
          .map(
            (width, i) => `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`,
          )
          .join('')}</cols>`
      : ''
  const dimension = `A1:${columnLabel(maxColumns - 1)}${Math.max(sheet.rows.length, 1)}`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="${dimension}"/>${cols}<sheetData>${lines.join('')}</sheetData></worksheet>`
  )
}

export async function buildTemplateXlsx(sheets: TemplateSheet[]): Promise<Buffer> {
  if (sheets.length === 0) throw new Error('A sheets template needs at least one sheet.')
  for (const sheet of sheets) validateSheetName(sheet.name)
  const zip = new JSZip()
  const overrides = sheets
    .map(
      (sheet, i) =>
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('')
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      overrides +
      '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  const sheetsXml = sheets
    .map(
      (sheet, i) =>
        `<sheet name="${escapeXml(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
    )
    .join('')
  const sheetRels = sheets
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    )
    .join('')
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets>${sheetsXml}</sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheetRels +
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
      `<Relationship Id="rId${sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/theme/theme1.xml', DEFAULT_THEME_XML)
  zip.file('xl/styles.xml', TEMPLATE_STYLESHEET_XML)
  sheets.forEach((sheet, i) => {
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, buildWorksheetXml(sheet))
  })
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const EMU_PER_INCH = 914400
const TITLE_BOX = {
  x: 0.5 * EMU_PER_INCH,
  y: 0.4 * EMU_PER_INCH,
  cx: 12.33 * EMU_PER_INCH,
  cy: 1.4 * EMU_PER_INCH,
}
const BODY_BOX = {
  x: 1 * EMU_PER_INCH,
  y: 2.1 * EMU_PER_INCH,
  cx: 11.33 * EMU_PER_INCH,
  cy: 4.9 * EMU_PER_INCH,
}
const TITLE_SLIDE_TITLE = {
  x: 0.5 * EMU_PER_INCH,
  y: 2.6 * EMU_PER_INCH,
  cx: 12.33 * EMU_PER_INCH,
  cy: 1.6 * EMU_PER_INCH,
}
const TITLE_SLIDE_SUB = {
  x: 0.5 * EMU_PER_INCH,
  y: 4.4 * EMU_PER_INCH,
  cx: 12.33 * EMU_PER_INCH,
  cy: 1.2 * EMU_PER_INCH,
}

export async function buildTemplatePptx(slides: TemplateSlide[]): Promise<Uint8Array> {
  if (slides.length === 0) throw new Error('A slides template needs at least one slide.')
  const opened = await openPptx(await createBlankPptx())
  // duplicateSlide inserts AFTER its source, so each new slide duplicates
  // the current last slide: duplicating slide 0 every time would reverse
  // the deck (and re-fill an already-filled slide). The copy's cleared
  // text boxes are empty shells; the fresh boxes below carry the content.
  let target = opened.deck.slides[0]!
  slides.forEach((slide, index) => {
    if (index > 0) {
      const copy = duplicateSlide(opened, opened.deck.slides.length - 1, {
        clearText: true,
      })
      if (!copy) throw new Error('Could not add a slide to the template deck.')
      target = copy
    }
    if (index === 0) {
      addElement(target, {
        kind: 'textbox',
        offset: TITLE_SLIDE_TITLE,
        paragraphs: [{ runs: [{ text: slide.title, bold: true, fontSize: 40 }] }],
      })
      if (slide.subtitle) {
        addElement(target, {
          kind: 'textbox',
          offset: TITLE_SLIDE_SUB,
          paragraphs: [{ runs: [{ text: slide.subtitle, fontSize: 20 }] }],
        })
      }
      if (slide.bullets.length > 0) {
        addElement(target, {
          kind: 'textbox',
          offset: {
            x: 1 * EMU_PER_INCH,
            y: 5.8 * EMU_PER_INCH,
            cx: 11.33 * EMU_PER_INCH,
            cy: 1.2 * EMU_PER_INCH,
          },
          paragraphs: slide.bullets.map((bullet) => ({
            runs: [{ text: `• ${bullet}`, fontSize: 18 }],
          })),
        })
      }
    } else {
      addElement(target, {
        kind: 'textbox',
        offset: TITLE_BOX,
        paragraphs: [{ runs: [{ text: slide.title, bold: true, fontSize: 32 }] }],
      })
      if (slide.bullets.length > 0) {
        addElement(target, {
          kind: 'textbox',
          offset: BODY_BOX,
          paragraphs: slide.bullets.map((bullet) => ({
            runs: [{ text: `• ${bullet}`, fontSize: 20 }],
          })),
        })
      }
    }
  })
  return savePptx(opened)
}
