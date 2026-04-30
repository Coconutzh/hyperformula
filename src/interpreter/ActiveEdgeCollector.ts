/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {SimpleCellAddress} from '../Cell'
import {FormulaVertex} from '../DependencyGraph/FormulaVertex'

export type ActiveDependency =
  | { kind: 'CELL', address: SimpleCellAddress }
  | { kind: 'RANGE', start: SimpleCellAddress, end: SimpleCellAddress }
  | { kind: 'RANGE_CELL', start: SimpleCellAddress, end: SimpleCellAddress, address: SimpleCellAddress }
  | { kind: 'NAMED_EXPRESSION', expressionName: string, address: SimpleCellAddress }

export interface ActiveEdgeSnapshot {
  byFormula: Map<number, ActiveDependency[]>,
  byDependency: Map<string, number[]>,
}

export class ActiveEdgeCollector {
  private readonly dependenciesByFormula = new Map<number, Set<string>>()
  private readonly dependenciesByKey = new Map<string, ActiveDependency>()

  public recordCellEdge(from: FormulaVertex | undefined, address: SimpleCellAddress): void {
    if (from?.idInGraph === undefined) {
      return
    }
    const dependency: ActiveDependency = { kind: 'CELL', address }
    this.recordDependency(from.idInGraph, dependency)
  }

  public recordRangeEdge(from: FormulaVertex | undefined, start: SimpleCellAddress, end: SimpleCellAddress): void {
    if (from?.idInGraph === undefined) {
      return
    }
    const dependency: ActiveDependency = { kind: 'RANGE', start, end }
    this.recordDependency(from.idInGraph, dependency)
  }

  public recordRangeCellEdge(from: FormulaVertex | undefined, start: SimpleCellAddress, end: SimpleCellAddress, address: SimpleCellAddress): void {
    if (from?.idInGraph === undefined) {
      return
    }
    const dependency: ActiveDependency = { kind: 'RANGE_CELL', start, end, address }
    this.recordDependency(from.idInGraph, dependency)
  }

  public recordNamedExpressionEdge(from: FormulaVertex | undefined, expressionName: string, address: SimpleCellAddress): void {
    if (from?.idInGraph === undefined) {
      return
    }
    const dependency: ActiveDependency = { kind: 'NAMED_EXPRESSION', expressionName, address }
    this.recordDependency(from.idInGraph, dependency)
  }

  public clearFormulaEdges(from: FormulaVertex | undefined): void {
    if (from?.idInGraph === undefined) {
      return
    }
    this.dependenciesByFormula.delete(from.idInGraph)
  }

  public snapshot(): ActiveEdgeSnapshot {
    const byFormula = new Map<number, ActiveDependency[]>()
    const byDependency = new Map<string, number[]>()

    for (const [formulaId, dependencyKeys] of this.dependenciesByFormula.entries()) {
      const dependencies: ActiveDependency[] = []
      for (const key of dependencyKeys.values()) {
        const dependency = this.dependenciesByKey.get(key)
        if (dependency !== undefined) {
          dependencies.push(dependency)
          const formulas = byDependency.get(key) ?? []
          formulas.push(formulaId)
          byDependency.set(key, formulas)
        }
      }
      byFormula.set(formulaId, dependencies)
    }

    return { byFormula, byDependency }
  }

  private recordDependency(formulaId: number, dependency: ActiveDependency): void {
    const key = ActiveEdgeCollector.toDependencyKey(dependency)
    this.dependenciesByKey.set(key, dependency)
    const formulaSet = this.dependenciesByFormula.get(formulaId) ?? new Set<string>()
    formulaSet.add(key)
    this.dependenciesByFormula.set(formulaId, formulaSet)
  }

  private static toDependencyKey(dependency: ActiveDependency): string {
    switch (dependency.kind) {
      case 'CELL':
        return `C:${dependency.address.sheet}:${dependency.address.col}:${dependency.address.row}`
      case 'RANGE':
        return `R:${dependency.start.sheet}:${dependency.start.col}:${dependency.start.row}:${dependency.end.col}:${dependency.end.row}`
      case 'RANGE_CELL':
        return `RC:${dependency.start.sheet}:${dependency.start.col}:${dependency.start.row}:${dependency.end.col}:${dependency.end.row}:${dependency.address.sheet}:${dependency.address.col}:${dependency.address.row}`
      case 'NAMED_EXPRESSION':
        return `N:${dependency.expressionName.toLowerCase()}:${dependency.address.sheet}:${dependency.address.col}:${dependency.address.row}`
    }
  }
}
