/**
 * Spreadsheet cell edits for Build chat: set literal values in an
 * existing .xlsx, in place, preserving the rest of the workbook.
 *
 * The edit is surgical string work on the target sheet's XML — the
 * same approach the xlsx gateway uses for its in-place edits — so
 * styles, other sheets, drawings and every other part ride through
 * untouched. A replaced cell keeps its style index; a string value is
 * written as an inline string and a number as a bare <v>, exactly the
 * shapes the app's own template writer emits. A cell that held a
 * formula ends up holding the literal value (that is what "set this
 * cell to X" means); the calculation chain is dropped when present,
 * since it may point at formulas that no longer exist and readers
 * rebuild it.
 */

import JSZip from 'jszip'

/** One literal cell assignment; `cell` is a normalized A1 reference. */
export interface SheetCellUpdate {
  cell: string
  value: string | number
}

function escapeXml(value: string): string {
  const escaped = value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
  let out = ''
  for (let i = 0; i < escaped.length; i++) {
    const code = escaped.charCodeAt(i)
    if (code <= 8 || code === 11 || code === 12 || (code >= 14 && code <= 31)) continue
    out += escaped[i]
  }
  return out
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

function splitRef(ref: string): { col: number; row: number } {
  const match = /^([A-Z]+)([0-9]+)$/.exec(ref)
  if (!match) throw new Error(`Bad cell reference: ${ref}`)
  let column = 0
  for (const ch of match[1]!) column = column * 26 + (ch.charCodeAt(0) - 64)
  return { col: column - 1, row: Number.parseInt(match[2]!, 10) }
}

function cellXml(ref: string, value: string | number, style: string | null): string {
  const styleAttr = style ? ` s="${style}"` : ''
  if (typeof value === 'number') return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`
  return (
    `<c r="${ref}"${styleAttr} t="inlineStr">` +
    `<is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
  )
}

interface Span {
  start: number
  end: number
  text: string
}

/** Every element with this tag in an XML string, as spans (self-closing
 *  or with a matching close tag; these parts never nest same-tag). */
function scanElements(xml: string, tag: string): Span[] {
  const pattern = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g')
  const spans: Span[] = []
  let match: RegExpExecArray | null
  while ((match = pattern.exec(xml)) !== null) {
    spans.push({ start: match.index, end: match.index + match[0].length, text: match[0] })
  }
  return spans
}

function attrOf(elementText: string, name: string): string | null {
  const openTag = elementText.slice(0, elementText.indexOf('>') + 1)
  const match = new RegExp(`\\b${name}="([^"]*)"`).exec(openTag)
  return match ? match[1]! : null
}

function setCellInSheetData(data: string, update: SheetCellUpdate): string {
  const { col, row } = splitRef(update.cell)
  const rows = scanElements(data, 'row')
  const rowSpan = rows.find((span) => attrOf(span.text, 'r') === String(row))
  if (!rowSpan) {
    const rowXml = `<row r="${row}">${cellXml(update.cell, update.value, null)}</row>`
    const after = rows.find((span) => Number(attrOf(span.text, 'r')) > row)
    if (after) return data.slice(0, after.start) + rowXml + data.slice(after.start)
    return data + rowXml
  }
  const rowText = rowSpan.text
  const selfClosing = rowText.endsWith('/>')
  const openTagEnd = rowText.indexOf('>') + 1
  const openTag = rowText.slice(0, openTagEnd)
  const inner = selfClosing ? '' : rowText.slice(openTagEnd, rowText.length - '</row>'.length)
  const cells = scanElements(inner, 'c')
  const existing = cells.find((span) => attrOf(span.text, 'r') === update.cell)
  let nextInner: string
  if (existing) {
    // a replaced cell keeps its style; its old value/formula goes away
    const style = attrOf(existing.text, 's')
    nextInner =
      inner.slice(0, existing.start) +
      cellXml(update.cell, update.value, style) +
      inner.slice(existing.end)
  } else {
    const insertAt = cells.find((span) => {
      const ref = attrOf(span.text, 'r')
      return ref !== null && splitRef(ref).col > col
    })
    const cell = cellXml(update.cell, update.value, null)
    nextInner = insertAt
      ? inner.slice(0, insertAt.start) + cell + inner.slice(insertAt.start)
      : inner + cell
  }
  const nextRow = selfClosing
    ? `${openTag.slice(0, -2)}>${nextInner}</row>`
    : `${openTag}${nextInner}</row>`
  return data.slice(0, rowSpan.start) + nextRow + data.slice(rowSpan.end)
}

/** Grow the sheet's declared dimension when an edit reaches past it —
 *  some readers bound their parse by the dimension. */
function expandDimension(xml: string, updates: readonly SheetCellUpdate[]): string {
  const dim = /<dimension\b[^>]*\bref="([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+))?"[^>]*\/>/.exec(xml)
  if (!dim) return xml
  const startRef = dim[1]!
  const end = splitRef(dim[2] ?? dim[1]!)
  let maxCol = end.col
  let maxRow = end.row
  for (const update of updates) {
    const part = splitRef(update.cell)
    if (part.col > maxCol) maxCol = part.col
    if (part.row > maxRow) maxRow = part.row
  }
  if (maxCol === end.col && maxRow === end.row) return xml
  const widened = dim[0].replace(/ref="[^"]*"/, `ref="${startRef}:${columnLabel(maxCol)}${maxRow}"`)
  return xml.replace(dim[0], widened)
}

/**
 * Apply literal cell updates to one sheet of an existing workbook and
 * return the new file bytes. Throws a plain-language error when the
 * file or the named sheet cannot be edited; the input buffer and every
 * untouched part are never modified.
 */
export async function applySheetCellUpdates(
  bytes: Buffer,
  updates: readonly SheetCellUpdate[],
  sheetName?: string,
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(bytes)
  const workbookEntry = zip.file('xl/workbook.xml')
  const relsEntry = zip.file('xl/_rels/workbook.xml.rels')
  if (!workbookEntry || !relsEntry) throw new Error('That file is not a readable spreadsheet.')
  const workbookXml = await workbookEntry.async('string')
  const relsXml = await relsEntry.async('string')

  const relTargets = new Map<string, string>()
  for (const span of scanElements(relsXml, 'Relationship')) {
    const id = attrOf(span.text, 'Id')
    const target = attrOf(span.text, 'Target')
    if (id && target) relTargets.set(id, target)
  }
  const sheets = scanElements(workbookXml, 'sheet').map((span) => ({
    name: attrOf(span.text, 'name'),
    rid: attrOf(span.text, 'r:id'),
  }))
  const chosen = sheetName ? sheets.find((sheet) => sheet.name === sheetName) : sheets[0]
  if (!chosen) {
    throw new Error(
      sheetName
        ? `There is no sheet named “${sheetName}” in that file.`
        : 'That workbook has no sheets.',
    )
  }
  const target = chosen.rid ? relTargets.get(chosen.rid) : undefined
  if (!target) throw new Error('That workbook is missing its sheet data.')
  const sheetPath = target.startsWith('/') ? target.slice(1) : `xl/${target}`
  const sheetEntry = zip.file(sheetPath)
  if (!sheetEntry) throw new Error('That workbook is missing its sheet data.')

  let xml = await sheetEntry.async('string')
  if (/<sheetData\s*\/>/.test(xml)) {
    xml = xml.replace(/<sheetData\s*\/>/, '<sheetData></sheetData>')
  }
  const openMatch = /<sheetData[^>]*>/.exec(xml)
  const closeIndex = xml.indexOf('</sheetData>')
  if (!openMatch || closeIndex < 0) throw new Error('That sheet has no editable cell area.')
  const dataStart = openMatch.index + openMatch[0].length
  let data = xml.slice(dataStart, closeIndex)
  for (const update of updates) data = setCellInSheetData(data, update)
  xml = xml.slice(0, dataStart) + data + xml.slice(closeIndex)
  xml = expandDimension(xml, updates)
  zip.file(sheetPath, xml)

  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml')
    const typesEntry = zip.file('[Content_Types].xml')
    if (typesEntry) {
      const typesXml = await typesEntry.async('string')
      zip.file('[Content_Types].xml', typesXml.replace(/<Override[^>]*calcChain\.xml"[^>]*\/>/, ''))
    }
    const cleanedRels = relsXml.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/, '')
    if (cleanedRels !== relsXml) zip.file('xl/_rels/workbook.xml.rels', cleanedRels)
  }
  return zip.generateAsync({ type: 'nodebuffer' })
}
