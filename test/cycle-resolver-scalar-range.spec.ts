import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Cycle resolver with scalar coercion from ranges', () => {
  it('should not report a cycle when a vertical range is coerced to the current row cell', () => {
    const hf = HyperFormula.buildFromArray([
      [null, '=B1:B2+1'],
      [null, null],
      [null, null],
    ], {licenseKey: 'gpl-v3'})

    hf.setCellContents(adr('B1'), 10)
    hf.setCellContents(adr('B2'), '=C1+1')
    hf.setCellContents(adr('C1'), '=B1+1')

    expect(hf.getCellValue(adr('C1'))).toBe(11)
    expect(hf.getCellValue(adr('B2'))).toBe(12)

    hf.destroy()
  })
})
