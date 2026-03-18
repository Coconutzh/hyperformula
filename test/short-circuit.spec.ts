import {FunctionArgumentType, FunctionPlugin, HyperFormula, ImplementedFunctions} from '../src'
import {adr} from './testUtils'

class ShortCircuitProbePlugin extends FunctionPlugin {
  public static implementedFunctions: ImplementedFunctions = {
    BUMP: {
      method: 'bump',
      parameters: [],
    },
  }

  public bump(ast: any, state: any): any {
    return this.runFunction(ast.args, state, this.metadata('BUMP'), () => {
      const context = (this.config.context ?? {}) as {counter?: number}
      context.counter = (context.counter ?? 0) + 1
      return context.counter
    })
  }
}

describe('Boolean short-circuit evaluation', () => {
  beforeAll(() => {
    if (HyperFormula.getFunctionPlugin('BUMP') === undefined) {
      HyperFormula.registerFunctionPlugin(ShortCircuitProbePlugin)
    }
  })

  afterAll(() => {
    if (HyperFormula.getFunctionPlugin('BUMP') !== undefined) {
      HyperFormula.unregisterFunction('BUMP')
    }
  })

  const evaluateSingleFormula = (formula: string, context: {counter: number}) => {
    const hf = HyperFormula.buildFromArray([[formula]], {
      licenseKey: 'gpl-v3',
      context,
    })
    const value = hf.getCellValue(adr('A1'))
    hf.destroy()
    return value
  }

  it('IF should evaluate only the active branch', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=IF(TRUE(), 1, BUMP())', context)

    expect(value).toBe(1)
    expect(context.counter).toBe(0)
  })

  it('IFS should stop on the first true condition', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=IFS(TRUE(), 1, BUMP(), 2)', context)

    expect(value).toBe(1)
    expect(context.counter).toBe(0)
  })

  it('SWITCH should evaluate only the matched branch', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=SWITCH(1, 1, 10, BUMP(), 20, 30)', context)

    expect(value).toBe(10)
    expect(context.counter).toBe(0)
  })

  it('CHOOSE should evaluate only the selected argument', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=CHOOSE(1, 10, BUMP(), 30)', context)

    expect(value).toBe(10)
    expect(context.counter).toBe(0)
  })

  it('IFERROR should not evaluate fallback when first argument is not an error', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=IFERROR(10, BUMP())', context)

    expect(value).toBe(10)
    expect(context.counter).toBe(0)
  })

  it('IFNA should not evaluate fallback when first argument is not #N/A', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=IFNA(10, BUMP())', context)

    expect(value).toBe(10)
    expect(context.counter).toBe(0)
  })

  it('AND should short-circuit after FALSE', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=AND(FALSE(), BUMP())', context)

    expect(value).toBe(false)
    expect(context.counter).toBe(0)
  })

  it('OR should short-circuit after TRUE', () => {
    const context = {counter: 0}
    const value = evaluateSingleFormula('=OR(TRUE(), BUMP())', context)

    expect(value).toBe(true)
    expect(context.counter).toBe(0)
  })

  it('AND/OR should avoid propagating dead-branch errors', () => {
    const andValue = evaluateSingleFormula('=AND(FALSE(), 1/0)', {counter: 0})
    const orValue = evaluateSingleFormula('=OR(TRUE(), 1/0)', {counter: 0})

    expect(andValue).toBe(false)
    expect(orValue).toBe(true)
  })
})
