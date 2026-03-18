/* eslint-disable no-console */
import fs from 'fs'
import path from 'path'
import {HyperFormula} from '../src'

type SheetPayload = {
  name: string,
  cells: Array<Array<string | number | boolean>>,
}

type FormulaCellPayload = {
  sheet: string,
  address: string,
  row: number,
  col: number,
}

type InputPayload = {
  sheets: SheetPayload[],
  formulaCells: FormulaCellPayload[],
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
  const inputPathArg = args.named.input ?? args.positional[0]
  if (!inputPathArg) {
    throw new Error('Missing --input path')
  }
  const outPathArg = args.named.out ?? args.positional[1] ?? ''

  const inputPath = path.resolve(inputPathArg)
  if (!fs.existsSync(inputPath)) {
    throw new Error(`Input file not found: ${inputPath}`)
  }

  const raw = fs.readFileSync(inputPath, 'utf-8').replace(/^\uFEFF/, '')
  const payload = JSON.parse(raw) as InputPayload
  const sheetsObject: Record<string, Array<Array<string | number | boolean>>> = {}
  let usedRows = 0
  let usedCols = 0
  payload.sheets.forEach((sheet) => {
    sheetsObject[sheet.name] = sheet.cells
    if (sheet.cells.length > usedRows) {
      usedRows = sheet.cells.length
    }
    for (const row of sheet.cells) {
      if (row.length > usedCols) {
        usedCols = row.length
      }
    }
  })

  for (const cell of payload.formulaCells) {
    if (cell.row + 1 > usedRows) {
      usedRows = cell.row + 1
    }
    if (cell.col + 1 > usedCols) {
      usedCols = cell.col + 1
    }
  }

  const hf = HyperFormula.buildFromSheets(sheetsObject, {
    licenseKey: 'gpl-v3',
    maxRows: Math.max(40000, usedRows, 1),
    maxColumns: Math.max(18278, usedCols, 1),
  })
  const cells = payload.formulaCells.map((cell) => {
    const sheetId = hf.getSheetId(cell.sheet)
    const value = hf.getCellValue({sheet: sheetId, row: cell.row, col: cell.col})
    return {
      sheet: cell.sheet,
      address: cell.address,
      hfValue: normalizeHFValue(value),
    }
  })
  hf.destroy()

  const result = {
    formulaCellCount: cells.length,
    generatedAt: new Date().toISOString(),
    cells,
  }

  if (outPathArg) {
    const outPath = path.resolve(outPathArg)
    fs.mkdirSync(path.dirname(outPath), {recursive: true})
    fs.writeFileSync(outPath, JSON.stringify(result, null, 2), 'utf-8')
  } else {
    process.stdout.write(JSON.stringify(result, null, 2))
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err)
  process.exit(1)
})
