import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Cycle resolver with lookup-specific active edges', () => {
  it('should not report a cycle when VLOOKUP touches only a safe row inside a larger range', () => {
    const hf = HyperFormula.buildFromArray([
      ['=VLOOKUP(1, A3:B4, 2, FALSE())'],
      [null],
      [1, 10],
      [2, '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A1'))).toBe(10)

    hf.destroy()
  })
})
