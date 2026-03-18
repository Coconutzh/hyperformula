#!/usr/bin/env node
/**
 * Evaluate all formula cells from an Excel workbook with HyperFormula.
 *
 * Usage:
 *   node script/hf_eval_workbook.js --excel assets/table_test.xlsm --out logs/hf_eval.json
 */

const fs = require('fs')
const path = require('path')
const ExcelJS = require('exceljs')

function parseArgs(argv) {
  const args = {}
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i]
    if (token.startsWith('--')) {
      const key = token.slice(2)
      const next = argv[i + 1]
      if (next && !next.startsWith('--')) {
        args[key] = next
        i++
      } else {
        args[key] = 'true'
      }
    }
  }
  return args
}

function colToLetters(col) {
  let n = col
  let result = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    result = String.fromCharCode(65 + rem) + result
    n = Math.floor((n - 1) / 26)
  }
  return result
}

function toA1(row, col) {
  return `${colToLetters(col)}${row}`
}

function toCellContent(cell) {
  if (cell.formula) {
    const formulaText = String(cell.formula).trim()
    return formulaText.startsWith('=') ? formulaText : `=${formulaText}`
  }

  const value = cell.value
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
      return value.richText.map((part) => part.text || '').join('')
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

function normalizeHFValue(value) {
  if (value == null) {
    return ''
  }
  if (typeof value === 'number') {
    return Number.isInteger(value) ? value : value
  }
  if (typeof value === 'string' || typeof value === 'boolean') {
    return value
  }
  if (value instanceof Date) {
    return value.toISOString()
  }
  if (typeof value === 'object') {
    if (value.type && value.value != null) {
      return { errorType: value.type, errorValue: value.value }
    }
    if (Array.isArray(value)) {
      return value
    }
    try {
      return JSON.parse(JSON.stringify(value))
    } catch (e) {
      return String(value)
    }
  }
  return String(value)
}

async function main() {
  const args = parseArgs(process.argv.slice(2))
  const excelPath = args.excel
  if (!excelPath) {
    throw new Error('Missing --excel path')
  }

  const outPath = args.out || ''
  const absoluteExcelPath = path.resolve(excelPath)
  if (!fs.existsSync(absoluteExcelPath)) {
    throw new Error(`Excel file not found: ${absoluteExcelPath}`)
  }

  // Prefer bundled CommonJS build, fallback to ts-node on source.
  let hfModule
  const cjsEntry = path.resolve(__dirname, '../commonjs/index.js')
  if (fs.existsSync(cjsEntry)) {
    hfModule = require(cjsEntry)
  } else {
    require('ts-node/register/transpile-only')
    hfModule = require(path.resolve(__dirname, '../src/index.ts'))
  }
  const HyperFormula =
    hfModule.HyperFormula
    || (hfModule.default && hfModule.default.HyperFormula)
    || hfModule.default
    || hfModule

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(absoluteExcelPath)

  const sheetsObject = {}
  const formulaCells = []

  workbook.worksheets.forEach((ws) => {
    let maxRow = 0
    let maxCol = 0
    const coords = new Map()

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
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

    const matrix = []
    for (let r = 1; r <= maxRow; r++) {
      const line = []
      for (let c = 1; c <= maxCol; c++) {
        line.push(coords.get(`${r}:${c}`) ?? '')
      }
      matrix.push(line)
    }
    sheetsObject[ws.name] = matrix
  })

  const hf = HyperFormula.buildFromSheets(sheetsObject, { licenseKey: 'gpl-v3' })
  const resultCells = formulaCells.map((item) => {
    const sheetId = hf.getSheetId(item.sheet)
    const value = hf.getCellValue({ sheet: sheetId, row: item.row, col: item.col })
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
    fs.mkdirSync(path.dirname(absoluteOutPath), { recursive: true })
    fs.writeFileSync(absoluteOutPath, JSON.stringify(result, null, 2), 'utf-8')
  } else {
    process.stdout.write(JSON.stringify(result, null, 2))
  }
}

main().catch((err) => {
  console.error(err && err.stack ? err.stack : err)
  process.exit(1)
})
