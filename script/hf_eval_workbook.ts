/* eslint-disable no-console */
import fs from 'fs'
import path from 'path'
import ExcelJS from 'exceljs'
import {HyperFormula} from '../src'

type FormulaCell = {
  sheet: string,
  address: string,
  row: number,
  col: number,
}

function parseArgs(argv: string[]): { named: Record<string, string>, positional: string[] } {
  const named: Record<string, string> = {}
  const positional: string[] = []
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i]
    if (token.startsWith('--')) {
      const key = token.slice(2)
      const next = argv[i + 1]
      if (next && !next.startsWith('--')) {
        named[key] = next
        i++
      } else {
        named[key] = 'true'
      }
    } else {
      positional.push(token)
    }
  }
  return {named, positional}
}

function colToLetters(col: number): string {
  let n = col
  let result = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    result = String.fromCharCode(65 + rem) + result
    n = Math.floor((n - 1) / 26)
  }
  return result
}

function toA1(row: number, col: number): string {
  return `${colToLetters(col)}${row}`
}

function toCellContent(cell: ExcelJS.Cell): string | number | boolean {
  if (cell.formula) {
    const formulaText = String(cell.formula).trim()
    return formulaText.startsWith('=') ? formulaText : `=${formulaText}`
  }

  const value = cell.value as any
  if (value == null) {
    return ''
  }
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  if (typeof value === 'object') {
    if (value.richText) {
      return value.richText.map((part: {text?: string}) => part.text || '').join('')
    }
    if (value.text != null) {
      return String(value.text)
    }
    if (value.result != null) {
      return value.result
    }
  }
  return String(value)
}

function normalizeHFValue(value: any): any {
  if (value == null) {
    return ''
  }
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  if (typeof value === 'object') {
    if (value.type && value.value != null) {
      return {errorType: value.type, errorValue: value.value}
    }
    try {
      return JSON.parse(JSON.stringify(value))
    } catch (e) {
      return String(value)
    }
  }
  return String(value)
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2))
  const excelPathArg = args.named.excel ?? args.positional[0]
  if (!excelPathArg) {
    throw new Error('Missing --excel path')
  }
  const outPath = args.named.out ?? args.positional[1] ?? ''
  const absoluteExcelPath = path.resolve(excelPathArg)
  if (!fs.existsSync(absoluteExcelPath)) {
    throw new Error(`Excel file not found: ${absoluteExcelPath}`)
  }

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(absoluteExcelPath)

  const sheetsObject: Record<string, Array<Array<string | number | boolean>>> = {}
  const formulaCells: FormulaCell[] = []

  workbook.worksheets.forEach((ws) => {
    let maxRow = 0
    let maxCol = 0
    const coords = new Map<string, string | number | boolean>()

    ws.eachRow({includeEmpty: false}, (row, rowNumber) => {
      row.eachCell({includeEmpty: false}, (cell, colNumber) => {
        maxRow = Math.max(maxRow, rowNumber)
        maxCol = Math.max(maxCol, colNumber)
        coords.set(`${rowNumber}:${colNumber}`, toCellContent(cell))
        if (cell.formula) {
          formulaCells.push({
            sheet: ws.name,
            address: toA1(rowNumber, colNumber),
            row: rowNumber - 1,
            col: colNumber - 1,
          })
        }
      })
    })

    const matrix: Array<Array<string | number | boolean>> = []
    for (let r = 1; r <= maxRow; r++) {
      const line: Array<string | number | boolean> = []
      for (let c = 1; c <= maxCol; c++) {
        line.push(coords.get(`${r}:${c}`) ?? '')
      }
      matrix.push(line)
    }
    sheetsObject[ws.name] = matrix
  })

  const hf = HyperFormula.buildFromSheets(sheetsObject, {licenseKey: 'gpl-v3'})
  const resultCells = formulaCells.map((item) => {
    const sheetId = hf.getSheetId(item.sheet)
    const value = hf.getCellValue({sheet: sheetId, row: item.row, col: item.col})
    return {
      sheet: item.sheet,
      address: item.address,
      hfValue: normalizeHFValue(value),
    }
  })
  hf.destroy()

  const result = {
    excelPath: absoluteExcelPath,
    formulaCellCount: resultCells.length,
    generatedAt: new Date().toISOString(),
    cells: resultCells,
  }

  if (outPath) {
    const absoluteOutPath = path.resolve(outPath)
    fs.mkdirSync(path.dirname(absoluteOutPath), {recursive: true})
    fs.writeFileSync(absoluteOutPath, JSON.stringify(result, null, 2), 'utf-8')
  } else {
    process.stdout.write(JSON.stringify(result, null, 2))
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err)
  process.exit(1)
})
