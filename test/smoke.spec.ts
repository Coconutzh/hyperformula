import {ErrorType, HyperFormula} from '../src'
import {adr} from './testUtils'

describe('HyperFormula', () => {
  it('should build engine from array and evaluate formulas', () => {
    const data = [
      [1, 2, 3],
      [4, 5, 6],
      ['=SUM(A1:C1)', '=SUM(A2:C2)', '=SUM(A1:C2)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A3'))).toBe(6)
    expect(hf.getCellValue(adr('B3'))).toBe(15)
    expect(hf.getCellValue(adr('C3'))).toBe(21)
    expect(hf.getSheetDimensions(0)).toEqual({width: 3, height: 3})

    hf.destroy()
  })

  it('should evaluate arithmetic and logical formulas', () => {
    const data = [
      [10, 20, 30],
      ['=A1+B1+C1', '=A1*B1', '=C1/A1'],
      ['=IF(A1>5, "big", "small")', '=AND(A1>0, B1>0)', '=OR(A1<0, B1>0)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A2'))).toBe(60)
    expect(hf.getCellValue(adr('B2'))).toBe(200)
    expect(hf.getCellValue(adr('C2'))).toBe(3)

    expect(hf.getCellValue(adr('A3'))).toBe('big')
    expect(hf.getCellValue(adr('B3'))).toBe(true)
    expect(hf.getCellValue(adr('C3'))).toBe(true)

    hf.destroy()
  })

  it('should handle common spreadsheet functions', () => {
    const data = [
      [1, 2, 3, 4, 5],
      ['=SUM(A1:E1)', '=AVERAGE(A1:E1)', '=MIN(A1:E1)', '=MAX(A1:E1)', '=COUNT(A1:E1)'],
      ['=CONCATENATE("Hello", " ", "World")', '=LEN("Test")', '=UPPER("hello")', '=LOWER("HELLO")', '=ABS(-5)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A2'))).toBe(15)
    expect(hf.getCellValue(adr('B2'))).toBe(3)
    expect(hf.getCellValue(adr('C2'))).toBe(1)
    expect(hf.getCellValue(adr('D2'))).toBe(5)
    expect(hf.getCellValue(adr('E2'))).toBe(5)

    expect(hf.getCellValue(adr('A3'))).toBe('Hello World')
    expect(hf.getCellValue(adr('B3'))).toBe(4)
    expect(hf.getCellValue(adr('C3'))).toBe('HELLO')
    expect(hf.getCellValue(adr('D3'))).toBe('hello')
    expect(hf.getCellValue(adr('E3'))).toBe(5)

    hf.destroy()
  })

  it('should add and remove rows with formula updates', () => {
    const data = [
      [1],
      [2],
      [3],
      ['=SUM(A1:A3)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A4'))).toBe(6)

    hf.addRows(0, [1, 1])
    hf.setCellContents(adr('A2'), 10)

    expect(hf.getCellValue(adr('A5'))).toBe(16)

    hf.removeRows(0, [1, 1])

    expect(hf.getCellValue(adr('A4'))).toBe(6)

    hf.destroy()
  })

  it('should allow VLOOKUP over IF array-constructed 2-column table', () => {
    const data = Array.from({length: 112}, () => Array(64).fill(null))

    // Row 109-112 (1-based), columns AV/BD/BF/BL.
    data[108][55] = 'K1'
    data[109][55] = 'K2'
    data[110][55] = 'K3'
    data[111][55] = 'K4'

    data[108][47] = 100
    data[109][47] = 200
    data[110][47] = 300
    data[111][47] = 400

    data[108][57] = 'K2'
    data[108][63] = '=VLOOKUP(BF109,IF({1,0},BD109:BD112,AV109:AV112),2,0)'

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})
    expect(hf.getCellValue({sheet: 0, row: 108, col: 63})).toBe(200)
    hf.destroy()
  })

  it('should keep TRUNC stable on decimal carry boundaries', () => {
    const data = Array.from({length: 19}, () => Array(2).fill(null))

    data[7][1] = 0

    for (let row = 8; row < 19; row++) {
      data[row][1] = `=TRUNC(B${row}+0.001,3)`
    }

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('B17'))).toBe(0.009)
    expect(hf.getCellValue(adr('B18'))).toBe(0.01)
    expect(hf.getCellValue(adr('B19'))).toBe(0.011)

    hf.destroy()
  })

  it('should accept bare TRUE/FALSE literals in formulas', () => {
    const data = [
      ['a', 42],
      ['=IF(FALSE,1,2)', '=IF(TRUE,1,2)'],
      ['a', '=VLOOKUP(A3,A1:B1,2,FALSE)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('B2'))).toBe(1)
    expect(hf.getCellValue(adr('B3'))).toBe(42)

    hf.destroy()
  })

  it('should treat formula-produced empty strings as #VALUE! in arithmetic but keep true blanks as zero', () => {
    const data = [
      ['=""', '=1+A1', '=IFERROR(1+A1,"")'],
      [null, '=1+A2', '=IFERROR(1+A2,"")'],
      ['=""', '=10-A3', '=A3*2'],
      ['=""', '=A4^2', '=IFERROR(A4^2,"")'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    const b1 = hf.getCellValue(adr('B1')) as any
    expect(b1.type).toBe(ErrorType.VALUE)
    expect(hf.getCellValue(adr('C1'))).toBe('')

    expect(hf.getCellValue(adr('B2'))).toBe(1)
    expect(hf.getCellValue(adr('C2'))).toBe(1)

    const b3 = hf.getCellValue(adr('B3')) as any
    const c3 = hf.getCellValue(adr('C3')) as any
    const b4 = hf.getCellValue(adr('B4')) as any
    expect(b3.type).toBe(ErrorType.VALUE)
    expect(c3.type).toBe(ErrorType.VALUE)
    expect(b4.type).toBe(ErrorType.VALUE)
    expect(hf.getCellValue(adr('C4'))).toBe('')

    hf.destroy()
  })

  it('should propagate errors through AND/OR instead of short-circuiting them away', () => {
    const data = [
      ['=NA()'],
      ['=AND(FALSE,A1)'],
      ['=AND(TRUE,A1)'],
      ['=OR(TRUE,A1)'],
      ['=OR(FALSE,A1)'],
      ['=IF(AND(FALSE,A1),1,0)'],
      ['=IF(AND(TRUE,A1),1,0)'],
      ['=IF(OR(TRUE,A1),1,0)'],
      ['=IF(OR(FALSE,A1),1,0)'],
    ]

    const hf = HyperFormula.buildFromArray(data, {licenseKey: 'gpl-v3'})

    for (const addr of ['A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9']) {
      const value = hf.getCellValue(adr(addr)) as any
      expect(value.type).toBe(ErrorType.NA)
    }

    hf.destroy()
  })

  it('should fold scalar-vs-range boolean comparisons inside OR/AND for workbook parity', () => {
    const hf = HyperFormula.buildFromSheets({
      Main: [
        ['双端面机械密封', null, '=OR(Main!A1="",Main!A1=Ref!B1:B2)'],
        [null, null, '=IF(OR(Main!A1="",Main!A1=Ref!B1:B2),0,1)'],
        [null, null, '=AND(Main!A1<>"",Main!A1=Ref!B1:B2)'],
      ],
      Ref: [
        [null, 1],
        [null, '无'],
      ],
    }, {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue({sheet: 0, row: 0, col: 2})).toBe(false)
    expect(hf.getCellValue({sheet: 0, row: 1, col: 2})).toBe(1)
    expect(hf.getCellValue({sheet: 0, row: 2, col: 2})).toBe(false)

    hf.destroy()
  })

  it('should allow MIN/MAX over IF-produced range filters for workbook parity', () => {
    const hf = HyperFormula.buildFromArray([
      [0, -1, 2, 3, null, '=MIN(IF(A1:D1>0,A1:D1))', '=MAX(IF(A1:D1>0,A1:D1))'],
      [null, null, null, null, null, '=IF(TRUE,MIN(IF(A1:D1>0,A1:D1)),1)', '=IF(TRUE,MAX(IF(A1:D1>0,A1:D1)),1)'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('F1'))).toBe(2)
    expect(hf.getCellValue(adr('G1'))).toBe(3)
    expect(hf.getCellValue(adr('F2'))).toBe(2)
    expect(hf.getCellValue(adr('G2'))).toBe(3)

    hf.destroy()
  })

  it('should match Excel error semantics for zero raised to non-positive powers', () => {
    const hf = HyperFormula.buildFromArray([
      ['=0^0', '=0^(-2)', '=0^(-1/3)', '=POWER(0,-2)', '=POWER(0,0)'],
    ], {licenseKey: 'gpl-v3'})

    const a1 = hf.getCellValue(adr('A1')) as any
    const b1 = hf.getCellValue(adr('B1')) as any
    const c1 = hf.getCellValue(adr('C1')) as any
    const d1 = hf.getCellValue(adr('D1')) as any
    const e1 = hf.getCellValue(adr('E1')) as any

    expect(a1.type).toBe(ErrorType.NUM)
    expect(b1.type).toBe(ErrorType.DIV_BY_ZERO)
    expect(c1.type).toBe(ErrorType.DIV_BY_ZERO)
    expect(d1.type).toBe(ErrorType.DIV_BY_ZERO)
    expect(e1.type).toBe(ErrorType.NUM)

    hf.destroy()
  })

  it('should preserve Excel TEXT semantics for formula empty strings and non-numeric scalars', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TEXT("","0.000")'],
      ['=""', '=TEXT(A2,"0.000")'],
      [null, '=TEXT(A3,"0.000")'],
      [null, '=TEXT(A4,"0.000")'],
      ['=TEXT("123","0.000")'],
      ['=TEXT(TRUE,"0.000")'],
      ['=TEXT(FALSE,"0.000")'],
      ['=TEXT("2026-01-31","yyyy-mm-dd")'],
      ['=TEXT("abc","0.000")'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A1'))).toBe('')
    expect(hf.getCellValue(adr('B2'))).toBe('')
    expect(hf.getCellValue(adr('B3'))).toBe('0.000')
    expect(hf.getCellValue(adr('B4'))).toBe('0.000')
    expect(hf.getCellValue(adr('A5'))).toBe('123.000')
    expect(hf.getCellValue(adr('A6'))).toBe('TRUE')
    expect(hf.getCellValue(adr('A7'))).toBe('FALSE')
    expect(hf.getCellValue(adr('A8'))).toBe('2026-01-31')
    expect(hf.getCellValue(adr('A9'))).toBe('abc')

    hf.destroy()
  })
})
