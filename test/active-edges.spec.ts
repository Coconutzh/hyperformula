import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Active edge sampling', () => {
  it('should record only active cell reference for IF(TRUE()) branch', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2, '=IF(TRUE(), A1, B1)']
    ], {licenseKey: 'gpl-v3'})

    const value = hf.getCellValue(adr('C1'))
    expect(value).toBe(1)

    const snapshot = hf.evaluator.activeEdgeSnapshot
    expect(snapshot).toBeDefined()

    const dependencyKeys = new Set(snapshot ? [...snapshot.byDependency.keys()] : [])
    expect(dependencyKeys.has('C:0:0:0')).toBe(true)
    expect(dependencyKeys.has('C:0:1:0')).toBe(false)

    hf.destroy()
  })

  it('should record only active cell reference for IF(FALSE()) branch', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2, '=IF(FALSE(), A1, B1)']
    ], {licenseKey: 'gpl-v3'})

    const value = hf.getCellValue(adr('C1'))
    expect(value).toBe(2)

    const snapshot = hf.evaluator.activeEdgeSnapshot
    expect(snapshot).toBeDefined()

    const dependencyKeys = new Set(snapshot ? [...snapshot.byDependency.keys()] : [])
    expect(dependencyKeys.has('C:0:0:0')).toBe(false)
    expect(dependencyKeys.has('C:0:1:0')).toBe(true)

    hf.destroy()
  })
})
