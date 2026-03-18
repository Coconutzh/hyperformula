import {ErrorType, HyperFormula, SimpleCellAddress} from '../src'
import {simpleCellAddress} from '../src/Cell'

const a = (col: number, row: number, sheet: number = 0): SimpleCellAddress => simpleCellAddress(sheet, col, row)

const expectCycle = (value: any) => {
  expect(typeof value).toBe('object')
  expect(value.type).toBe(ErrorType.CYCLE)
}

describe('Cycle resolver - complex scenarios', () => {
  it('should resolve and re-detect a 3-node cycle guarded by IF during incremental updates', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IF(FALSE(),C1,10)', '=A1+1', '=B1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(10)
    expect(hf.getCellValue(a(1, 0))).toBe(11)
    expect(hf.getCellValue(a(2, 0))).toBe(12)

    hf.setCellContents(a(0, 0), '=IF(TRUE(),C1,10)')
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))
    expectCycle(hf.getCellValue(a(2, 0)))

    hf.setCellContents(a(0, 0), '=IF(FALSE(),C1,10)')
    expect(hf.getCellValue(a(0, 0))).toBe(10)
    expect(hf.getCellValue(a(1, 0))).toBe(11)
    expect(hf.getCellValue(a(2, 0))).toBe(12)

    hf.destroy()
  })

  it('should honor IFS short-circuit when deciding whether a cycle is active', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IFS(TRUE(),5,TRUE(),B1)', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(5)
    expect(hf.getCellValue(a(1, 0))).toBe(6)

    hf.setCellContents(a(0, 0), '=IFS(FALSE(),5,TRUE(),B1)')
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))

    hf.setCellContents(a(0, 0), '=IFS(TRUE(),5,TRUE(),B1)')
    expect(hf.getCellValue(a(0, 0))).toBe(5)
    expect(hf.getCellValue(a(1, 0))).toBe(6)

    hf.destroy()
  })

  it('should resolve cycle activation controlled by SWITCH selector cell', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SWITCH(D1,1,7,2,B1,0)', '=A1+1', '', 1],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(7)
    expect(hf.getCellValue(a(1, 0))).toBe(8)

    hf.setCellContents(a(3, 0), 2)
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))

    hf.setCellContents(a(3, 0), 1)
    expect(hf.getCellValue(a(0, 0))).toBe(7)
    expect(hf.getCellValue(a(1, 0))).toBe(8)

    hf.destroy()
  })

  it('should resolve cycle activation controlled by CHOOSE selector cell', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSE(D1,9,B1)', '=A1+1', '', 1],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(9)
    expect(hf.getCellValue(a(1, 0))).toBe(10)

    hf.setCellContents(a(3, 0), 2)
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))

    hf.setCellContents(a(3, 0), 1)
    expect(hf.getCellValue(a(0, 0))).toBe(9)
    expect(hf.getCellValue(a(1, 0))).toBe(10)

    hf.destroy()
  })

  it('should treat IFERROR fallback branch as active only when needed', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IFERROR(1,B1)', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(1)
    expect(hf.getCellValue(a(1, 0))).toBe(2)

    hf.setCellContents(a(0, 0), '=IFERROR(1/0,B1)')
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))

    hf.setCellContents(a(0, 0), '=IFERROR(1,B1)')
    expect(hf.getCellValue(a(0, 0))).toBe(1)
    expect(hf.getCellValue(a(1, 0))).toBe(2)

    hf.destroy()
  })

  it('should honor AND short-circuit in cycle activation', () => {
    const hf = HyperFormula.buildFromArray([
      ['=AND(FALSE(),B1>0)', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(false)
    expect(hf.getCellValue(a(1, 0))).toBe(1)

    hf.setCellContents(a(0, 0), '=AND(TRUE(),B1>0)')
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))

    hf.destroy()
  })

  it('should handle nested guard functions when deciding cycle activation', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IF(FALSE(),B1,IFERROR(1,C1))', '=A1+1', '=B1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0))).toBe(1)
    expect(hf.getCellValue(a(1, 0))).toBe(2)
    expect(hf.getCellValue(a(2, 0))).toBe(3)

    hf.setCellContents(a(0, 0), '=IF(TRUE(),B1,IFERROR(1,C1))')
    expectCycle(hf.getCellValue(a(0, 0)))
    expectCycle(hf.getCellValue(a(1, 0)))
    expectCycle(hf.getCellValue(a(2, 0)))

    hf.destroy()
  })

  it('should resolve cross-sheet cycle based on guard value on another sheet', () => {
    const hf = HyperFormula.buildFromSheets({
      Sheet1: [['=IF(Sheet2!A1=0,1,B1)', '=A1+1']],
      Sheet2: [[0]],
    }, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(a(0, 0, 0))).toBe(1)
    expect(hf.getCellValue(a(1, 0, 0))).toBe(2)

    hf.setCellContents(a(0, 0, 1), 1)
    expectCycle(hf.getCellValue(a(0, 0, 0)))
    expectCycle(hf.getCellValue(a(1, 0, 0)))

    hf.setCellContents(a(0, 0, 1), 0)
    expect(hf.getCellValue(a(0, 0, 0))).toBe(1)
    expect(hf.getCellValue(a(1, 0, 0))).toBe(2)

    hf.destroy()
  })
})
