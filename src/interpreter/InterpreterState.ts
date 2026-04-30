/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType, SimpleCellAddress} from '../Cell'
import {FormulaVertex} from '../DependencyGraph/FormulaVertex'
import {DependencyGraph} from '../DependencyGraph'
import {ErrorMessage} from '../error-message'
import {InternalScalarValue, InterpreterValue} from './InterpreterValue'
import {SimpleRangeValue} from '../SimpleRangeValue'
import {ActiveEdgeCollector} from './ActiveEdgeCollector'

export type ProbeAccess =
  | { kind: 'CELL', address: SimpleCellAddress }
  | { kind: 'RANGE', start: SimpleCellAddress, end: SimpleCellAddress }
  | { kind: 'RANGE_CELL', start: SimpleCellAddress, end: SimpleCellAddress, address: SimpleCellAddress }
  | { kind: 'NAMED_EXPRESSION', expressionName: string, address: SimpleCellAddress }

export class ProbeAccessTracker {
  private lastAccess?: ProbeAccess

  public recordCell(address: SimpleCellAddress): void {
    this.lastAccess = { kind: 'CELL', address }
  }

  public recordRange(start: SimpleCellAddress, end: SimpleCellAddress): void {
    this.lastAccess = { kind: 'RANGE', start, end }
  }

  public recordRangeCell(start: SimpleCellAddress, end: SimpleCellAddress, address: SimpleCellAddress): void {
    this.lastAccess = { kind: 'RANGE_CELL', start, end, address }
  }

  public recordNamedExpression(expressionName: string, address: SimpleCellAddress): void {
    this.lastAccess = { kind: 'NAMED_EXPRESSION', expressionName, address }
  }

  public snapshot(): ProbeAccess | undefined {
    return this.lastAccess
  }
}

export interface ProbeValueResolver {
  getCellValue(address: SimpleCellAddress): InterpreterValue,
}

export class InterpreterState {
  constructor(
    public formulaAddress: SimpleCellAddress,
    public arraysFlag: boolean,
    public formulaVertex?: FormulaVertex,
    public activeEdgeCollector?: ActiveEdgeCollector,
    public probeAccessTracker?: ProbeAccessTracker,
    public probeValueResolver?: ProbeValueResolver,
  ) {
  }

  public recordCellAccess(address: SimpleCellAddress): void {
    this.activeEdgeCollector?.recordCellEdge(this.formulaVertex, address)
    this.probeAccessTracker?.recordCell(address)
  }

  public recordRangeAccess(start: SimpleCellAddress, end: SimpleCellAddress): void {
    this.activeEdgeCollector?.recordRangeEdge(this.formulaVertex, start, end)
    this.probeAccessTracker?.recordRange(start, end)
  }

  public recordRangeCellAccess(start: SimpleCellAddress, end: SimpleCellAddress, address: SimpleCellAddress): void {
    this.activeEdgeCollector?.recordRangeCellEdge(this.formulaVertex, start, end, address)
    this.probeAccessTracker?.recordRangeCell(start, end, address)
  }

  public recordNamedExpressionAccess(expressionName: string, address: SimpleCellAddress): void {
    this.activeEdgeCollector?.recordNamedExpressionEdge(this.formulaVertex, expressionName, address)
    this.probeAccessTracker?.recordNamedExpression(expressionName, address)
  }

  public shouldTrackRangeAccess(): boolean {
    return this.formulaVertex !== undefined || this.activeEdgeCollector !== undefined || this.probeAccessTracker !== undefined
  }

  public getCellValue(dependencyGraph: DependencyGraph, address: SimpleCellAddress): InterpreterValue {
    return this.probeValueResolver?.getCellValue(address) ?? dependencyGraph.getCellValue(address)
  }

  public getScalarValue(dependencyGraph: DependencyGraph, address: SimpleCellAddress): InternalScalarValue {
    const value = this.getCellValue(dependencyGraph, address)
    if (value instanceof SimpleRangeValue) {
      return new CellError(ErrorType.VALUE, ErrorMessage.ScalarExpected)
    }
    return value
  }
}
