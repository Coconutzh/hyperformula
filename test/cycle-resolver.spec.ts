import {ErrorType, HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Cycle resolver with active edges', () => {
  it('should resolve static cycle when active branch cuts dependency', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IF(FALSE(), B1, 1)', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)

    hf.destroy()
  })

  it('should keep reporting cycle when active graph is still cyclic', () => {
    const hf = HyperFormula.buildFromArray([
      ['=B1+1', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    const a1 = hf.getCellValue(adr('A1'))
    const b1 = hf.getCellValue(adr('B1'))

    expect(typeof a1).toBe('object')
    expect((a1 as any).type).toBe(ErrorType.CYCLE)
    expect(typeof b1).toBe('object')
    expect((b1 as any).type).toBe(ErrorType.CYCLE)

    hf.destroy()
  })
})
