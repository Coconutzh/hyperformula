import {HyperFormula} from '../src'

describe('Cycle resolver probe debug snapshot', () => {
  it('should capture the last accessed cell for probe failures', () => {
    const hf = HyperFormula.buildFromArray([
      ['=IF(TRUE(),B1,1)', '=A1+1'],
    ], {licenseKey: 'gpl-v3'})

    const snapshot = hf.getCycleResolutionSnapshot()
    expect(snapshot.unresolvedSccCount).toBe(1)

    const scc = snapshot.unresolvedSccs[0]
    const formula = scc.formulas.find((candidate) => /A1$/.test(candidate.address))
    expect(formula).toBeDefined()
    expect(formula?.probeFailed).toBe(true)
    expect(formula?.probeFailure?.errorType).toBe('Error')
    expect(formula?.probeFailure?.errorMessage).toContain('not computed')
    expect(formula?.probeFailure?.lastAccessKind).toBe('CELL')
    expect(formula?.probeFailure?.lastAccessAddress).toMatch(/B1$/)

    hf.destroy()
  })

  it('should capture the last accessed range cell and range for probe failures', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MAX(B1:B2)', '=A1+1'],
      ['', 2],
    ], {licenseKey: 'gpl-v3'})

    const snapshot = hf.getCycleResolutionSnapshot()
    expect(snapshot.unresolvedSccCount).toBe(1)

    const scc = snapshot.unresolvedSccs[0]
    const formula = scc.formulas.find((candidate) => /A1$/.test(candidate.address))
    expect(formula).toBeDefined()
    expect(formula?.probeFailed).toBe(true)
    expect(formula?.probeFailure?.lastAccessKind).toBe('RANGE_CELL')
    expect(formula?.probeFailure?.lastAccessAddress).toMatch(/B1$/)
    expect(formula?.probeFailure?.lastAccessRange).toMatch(/B1:B2$/)

    hf.destroy()
  })
})
