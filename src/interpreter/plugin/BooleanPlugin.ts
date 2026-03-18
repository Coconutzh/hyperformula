/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../Cell'
import {ErrorMessage} from '../../error-message'
import {ProcedureAst} from '../../parser'
import {InterpreterState} from '../InterpreterState'
import {InternalNoErrorScalarValue, InternalScalarValue, InterpreterValue} from '../InterpreterValue'
import {FunctionArgument, FunctionArgumentType, FunctionMetadata, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from './FunctionPlugin'

/**
 * Interpreter plugin containing boolean functions
 */
export class BooleanPlugin extends FunctionPlugin implements FunctionPluginTypecheck<BooleanPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'TRUE': {
      method: 'literalTrue',
      parameters: [],
    },
    'FALSE': {
      method: 'literalFalse',
      parameters: [],
    },
    'IF': {
      method: 'conditionalIf',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
        {argumentType: FunctionArgumentType.SCALAR, defaultValue: false, passSubtype: true},
      ],
    },
    'IFS': {
      method: 'ifs',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
      ],
      repeatLastArgs: 2,
    },
    'AND': {
      method: 'and',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'OR': {
      method: 'or',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'XOR': {
      method: 'xor',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'NOT': {
      method: 'not',
      parameters: [
        {argumentType: FunctionArgumentType.BOOLEAN},
      ]
    },
    'SWITCH': {
      method: 'switch',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
      ],
      repeatLastArgs: 1,
    },
    'IFERROR': {
      method: 'iferror',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
      ]
    },
    'IFNA': {
      method: 'ifna',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
      ]
    },
    'CHOOSE': {
      method: 'choose',
      parameters: [
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1},
        {argumentType: FunctionArgumentType.SCALAR, passSubtype: true},
      ],
      repeatLastArgs: 1,
    },
  }

  /**
   * Corresponds to TRUE()
   *
   * Returns the logical true
   *
   * @param ast
   * @param state
   */
  public literalTrue(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TRUE'), () => true)
  }

  /**
   * Corresponds to FALSE()
   *
   * Returns the logical false
   *
   * @param ast
   * @param state
   */
  public literalFalse(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('FALSE'), () => false)
  }

  /**
   * Corresponds to IF(expression, value_if_true, value_if_false)
   *
   * Returns value specified as second argument if expression is true and third argument if expression is false
   *
   * @param ast
   * @param state
   */
  public conditionalIf(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('IF')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const condition = this.evaluateAstAsType(ast.args[0], state, metadata.parameters![0]) as boolean | CellError | undefined
    if (condition === undefined) {
      return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
    }
    if (condition instanceof CellError) {
      return condition
    }

    if (condition) {
      return this.evaluateAst(ast.args[1], state)
    }

    if (ast.args[2] !== undefined) {
      return this.evaluateAst(ast.args[2], state)
    }

    return metadata.parameters![2].defaultValue as InternalScalarValue
  }

  /**
   * Implementation for the IFS function. Returns the value that corresponds to the first true condition.
   *
   * @param ast
   * @param state
   */
  public ifs(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('IFS')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    for (let idx = 0; idx < ast.args.length; idx += 2) {
      const condition = this.evaluateAstAsType(ast.args[idx], state, metadata.parameters![0]) as boolean | CellError | undefined
      if (condition === undefined) {
        return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
      }
      if (condition instanceof CellError) {
        return condition
      }
      if (condition) {
        return this.evaluateAst(ast.args[idx + 1], state)
      }
    }

    return new CellError(ErrorType.NA, ErrorMessage.NoConditionMet)
  }

  /**
   * Corresponds to AND(expression1, [expression2, ...])
   *
   * Returns true if all of the provided arguments are logically true, and false if any of it is logically false
   *
   * @param ast
   * @param state
   */
  public and(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('AND')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const booleanMeta: FunctionArgument = {argumentType: FunctionArgumentType.BOOLEAN}
    for (const arg of ast.args) {
      const scalarValues = this.listOfScalarValues([arg], state)
      for (const [scalarValue] of scalarValues) {
        const coerced = this.coerceToType(scalarValue, booleanMeta, state)
        if (coerced instanceof CellError) {
          return coerced
        }
        if (coerced !== undefined && !coerced) {
          return false
        }
      }
    }

    return true
  }

  /**
   * Corresponds to OR(expression1, [expression2, ...])
   *
   * Returns true if any of the provided arguments are logically true, and false otherwise
   *
   * @param ast
   * @param state
   */
  public or(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('OR')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const booleanMeta: FunctionArgument = {argumentType: FunctionArgumentType.BOOLEAN}
    for (const arg of ast.args) {
      const scalarValues = this.listOfScalarValues([arg], state)
      for (const [scalarValue] of scalarValues) {
        const coerced = this.coerceToType(scalarValue, booleanMeta, state)
        if (coerced instanceof CellError) {
          return coerced
        }
        if (coerced !== undefined && coerced) {
          return true
        }
      }
    }

    return false
  }

  public not(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('NOT'), (arg) => !arg)
  }

  public xor(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('XOR'), (...args: (boolean | undefined)[]) => {
      let cnt = 0
      args.filter(arg => arg !== undefined).forEach(arg => {
        if (arg) {
          cnt++
        }
      })
      return (cnt % 2) === 1
    })
  }

  public switch(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('SWITCH')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const selector = this.evaluateAstAsType(ast.args[0], state, metadata.parameters![0]) as InternalScalarValue | CellError | undefined
    if (selector instanceof CellError) {
      return selector
    }
    if (selector === undefined) {
      return new CellError(ErrorType.VALUE, ErrorMessage.CellRefExpected)
    }

    let idx = 1
    while (idx + 1 < ast.args.length) {
      const caseValue = this.evaluateAstAsType(ast.args[idx], state, metadata.parameters![1]) as InternalScalarValue | CellError | undefined
      if (!(caseValue instanceof CellError) && caseValue !== undefined && this.arithmeticHelper.eq(selector, caseValue as InternalNoErrorScalarValue)) {
        return this.evaluateAst(ast.args[idx + 1], state)
      }
      idx += 2
    }

    if (idx < ast.args.length) {
      return this.evaluateAst(ast.args[idx], state)
    }

    return new CellError(ErrorType.NA, ErrorMessage.NoDefault)
  }

  public iferror(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('IFERROR')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const arg1 = this.evaluateAst(ast.args[0], state)
    if (arg1 instanceof CellError) {
      return this.evaluateAst(ast.args[1], state)
    }
    return arg1
  }

  public ifna(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('IFNA')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const arg1 = this.evaluateAst(ast.args[0], state)
    if (arg1 instanceof CellError && arg1.type === ErrorType.NA) {
      return this.evaluateAst(ast.args[1], state)
    }
    return arg1
  }

  public choose(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const metadata = this.metadata('CHOOSE')
    const validArgNumber = this.validateNumberOfArguments(ast.args.length, metadata)
    if (validArgNumber !== undefined) {
      return validArgNumber
    }

    const selector = this.evaluateAstAsType(ast.args[0], state, metadata.parameters![0]) as number | CellError | undefined
    if (selector === undefined) {
      return new CellError(ErrorType.NUM, ErrorMessage.Selector)
    }
    if (selector instanceof CellError) {
      return selector
    }

    const selectedArgIndex = selector
    const numberOfChoices = ast.args.length - 1
    if (selectedArgIndex > numberOfChoices) {
      return new CellError(ErrorType.NUM, ErrorMessage.Selector)
    }

    return this.evaluateAst(ast.args[selectedArgIndex], state)
  }

  private validateNumberOfArguments(numberOfArgumentsPassed: number, metadata: FunctionMetadata): CellError | undefined {
    const argumentsMetadata = this.buildMetadataForEachArgumentValue(numberOfArgumentsPassed, metadata)
    if (!this.isNumberOfArgumentValuesValid(argumentsMetadata, numberOfArgumentsPassed)) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }
    return undefined
  }

  private evaluateAstAsType(ast: ProcedureAst['args'][number], state: InterpreterState, argumentType: FunctionArgument): InterpreterValue | boolean | number | string | undefined {
    const value = this.evaluateAst(ast, state)
    const coercedValue = this.coerceToType(value, argumentType, state)
    if (coercedValue === undefined) {
      return undefined
    }
    if (coercedValue instanceof CellError) {
      return coercedValue
    }
    return coercedValue as InterpreterValue | boolean | number | string
  }
}
