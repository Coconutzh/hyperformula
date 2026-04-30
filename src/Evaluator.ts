/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {AbsoluteCellRange} from './AbsoluteCellRange'
import {absolutizeDependencies} from './absolutizeDependencies'
import {CellError, ErrorType, SimpleCellAddress} from './Cell'
import {Config} from './Config'
import {ContentChanges} from './ContentChanges'
import {ArrayFormulaVertex, DependencyGraph, RangeVertex, Vertex} from './DependencyGraph'
import {FormulaVertex} from './DependencyGraph/FormulaVertex'
import {ActiveDependency, ActiveEdgeCollector, ActiveEdgeSnapshot} from './interpreter/ActiveEdgeCollector'
import {Interpreter} from './interpreter/Interpreter'
import {InterpreterState, ProbeAccess, ProbeAccessTracker, ProbeValueResolver} from './interpreter/InterpreterState'
import {EmptyValue, getRawValue, InterpreterValue} from './interpreter/InterpreterValue'
import {SimpleRangeValue} from './SimpleRangeValue'
import {LazilyTransformingAstService} from './LazilyTransformingAstService'
import {ColumnSearchStrategy} from './Lookup/SearchStrategy'
import {Ast, RelativeDependency, simpleCellAddressToString, simpleCellRangeToString, Unparser} from './parser'
import {Statistics, StatType} from './statistics'
import {TopSortResult} from './DependencyGraph/TopSort'

export interface CycleResolutionDebugDependency {
  kind: ActiveDependency['kind'],
  address?: string,
  range?: string,
  expressionName?: string,
  coveredByPreciseRange?: boolean,
}

export interface CycleResolutionDebugProbeFailure {
  errorType: string,
  errorMessage: string,
  lastAccessKind: ProbeAccess['kind'] | null,
  lastAccessAddress?: string,
  lastAccessRange?: string,
  expressionName?: string,
}

export interface CycleResolutionDebugFormula {
  formulaId: number | null,
  sheet: string,
  address: string,
  formula: string,
  status: 'ordered' | 'unresolved',
  probeFailed: boolean,
  probeFailure?: CycleResolutionDebugProbeFailure,
  dependencyCounts: {
    cell: number,
    range: number,
    rangeCell: number,
    namedExpression: number,
    coarseRange: number,
  },
  sccDependencyFormulaIds: number[],
  sccDependencyAddresses: string[],
  activeDependencies: CycleResolutionDebugDependency[],
}

export interface CycleResolutionDebugScc {
  sccIndex: number,
  formulaCount: number,
  orderedCount: number,
  unresolvedCount: number,
  unresolvedFormulaIds: number[],
  unresolvedAddresses: string[],
  probeFailedFormulaIds: number[],
  dependencyCounts: {
    cell: number,
    range: number,
    rangeCell: number,
    namedExpression: number,
    coarseRange: number,
  },
  reason: 'probe_failed_or_still_cyclic' | 'still_cyclic_after_active_edges',
  formulas: CycleResolutionDebugFormula[],
}

export interface CycleResolutionDebugSnapshot {
  totalCyclicSccCount: number,
  unresolvedSccCount: number,
  unresolvedFormulaCount: number,
  unresolvedSccs: CycleResolutionDebugScc[],
}

export class Evaluator {
  private activeEdgeCollector?: ActiveEdgeCollector
  private _activeEdgeSnapshot?: ActiveEdgeSnapshot
  private _cycleResolutionSnapshot: CycleResolutionDebugSnapshot = Evaluator.emptyCycleResolutionSnapshot()

  constructor(
    private readonly config: Config,
    private readonly stats: Statistics,
    public readonly interpreter: Interpreter,
    private readonly lazilyTransformingAstService: LazilyTransformingAstService,
    private readonly dependencyGraph: DependencyGraph,
    private readonly columnSearch: ColumnSearchStrategy,
    private readonly unparser: Unparser,
  ) {
  }

  public get activeEdgeSnapshot(): ActiveEdgeSnapshot | undefined {
    return this._activeEdgeSnapshot
  }

  public get cycleResolutionSnapshot(): CycleResolutionDebugSnapshot {
    return this._cycleResolutionSnapshot
  }

  public run(): void {
    this.activeEdgeCollector = new ActiveEdgeCollector()
    this.stats.start(StatType.TOP_SORT)
    const topSortResult = this.dependencyGraph.topSortWithScc()
    this.stats.end(StatType.TOP_SORT)

    this.stats.measure(StatType.EVALUATION, () => {
      this.recomputeFormulas(topSortResult)
    })
    this._activeEdgeSnapshot = this.activeEdgeCollector.snapshot()
    this.activeEdgeCollector = undefined
  }

  public partialRun(vertices: Vertex[]): ContentChanges {
    this.activeEdgeCollector = new ActiveEdgeCollector()
    const changes = ContentChanges.empty()

    this.stats.measure(StatType.EVALUATION, () => {
      const topSortResult = this.dependencyGraph.graph.getTopSortedWithSccSubgraphFrom(
        vertices,
        () => true,
        () => {},
      )
      this.recomputeFormulas(topSortResult, changes)
    })
    this._activeEdgeSnapshot = this.activeEdgeCollector.snapshot()
    this.activeEdgeCollector = undefined
    return changes
  }

  public runAndForget(ast: Ast, address: SimpleCellAddress, dependencies: RelativeDependency[]): InterpreterValue {
    const tmpRanges: RangeVertex[] = []
    for (const dep of absolutizeDependencies(dependencies, address)) {
      if (dep instanceof AbsoluteCellRange) {
        const range = dep
        if (this.dependencyGraph.getRange(range.start, range.end) === undefined) {
          const rangeVertex = new RangeVertex(range)
          this.dependencyGraph.rangeMapping.addOrUpdateVertex(rangeVertex)
          tmpRanges.push(rangeVertex)
        }
      }
    }
    const ret = this.evaluateAstToCellValue(ast, new InterpreterState(address, this.config.useArrayArithmetic))

    tmpRanges.forEach((rangeVertex) => {
      this.dependencyGraph.rangeMapping.removeVertexIfExists(rangeVertex)
    })

    return ret
  }

  /**
   * Recalculates the value of a single vertex assuming its dependencies have already been recalculated
   */
  private recomputeVertex(vertex: Vertex, changes: ContentChanges): boolean {
    if (vertex instanceof FormulaVertex) {
      const currentValue = vertex.isComputed() ? vertex.getCellValue() : undefined
      const newCellValue = this.recomputeFormulaVertexValue(vertex)
      if (newCellValue !== currentValue) {
        const address = vertex.getAddress(this.lazilyTransformingAstService)
        changes.addChange(newCellValue, address)
        this.columnSearch.change(getRawValue(currentValue), getRawValue(newCellValue), address)
        return true
      }
      return false
    } else if (vertex instanceof RangeVertex) {
      vertex.clearCache()
      return true
    } else {
      return true
    }
  }

  /**
   * Processes a vertex that is part of a cycle in dependency graph
   */
  private processVertexOnCycle(vertex: Vertex, changes: ContentChanges): void {
    if (vertex instanceof RangeVertex) {
      vertex.clearCache()
    } else if (vertex instanceof FormulaVertex) {
      const address = vertex.getAddress(this.lazilyTransformingAstService)
      this.columnSearch.remove(getRawValue(vertex.valueOrUndef()), address)
      const error = new CellError(ErrorType.CYCLE, undefined, vertex)
      vertex.setCellValue(error)
      changes.addChange(error, address)
    }
  }

  /**
   * Recalculates formulas in the topological sort order
   */
  private recomputeFormulas(topSortResult: TopSortResult<Vertex>, changes?: ContentChanges): void {
    const {sorted, cycled, cyclicSccs} = topSortResult
    const deferredFromSorted: FormulaVertex[] = []
    sorted.forEach((vertex: Vertex) => {
      if (changes !== undefined) {
        this.recomputeVertex(vertex, changes)
      } else if (vertex instanceof FormulaVertex) {
        try {
          const newCellValue = this.recomputeFormulaVertexValue(vertex)
          const address = vertex.getAddress(this.lazilyTransformingAstService)
          this.columnSearch.add(getRawValue(newCellValue), address)
        } catch (e) {
          deferredFromSorted.push(vertex)
        }
      } else if (vertex instanceof RangeVertex) {
        vertex.clearCache()
      }
    })

    this.resolveCyclicSccs(cycled, cyclicSccs, changes)

    if (changes === undefined) {
      deferredFromSorted.forEach((vertex) => {
        try {
          const newCellValue = this.recomputeFormulaVertexValue(vertex)
          const address = vertex.getAddress(this.lazilyTransformingAstService)
          this.columnSearch.add(getRawValue(newCellValue), address)
        } catch (e) {
          vertex.setCellValue(new CellError(ErrorType.CYCLE, undefined, vertex))
        }
      })
    }
  }

  private resolveCyclicSccs(cycled: Vertex[], cyclicSccs: Vertex[][], changes?: ContentChanges): void {
    const remainingCycledFormulas = new Set<FormulaVertex>(cycled.filter((vertex): vertex is FormulaVertex => vertex instanceof FormulaVertex))
    const processedFormulaIds = new Set<number>()
    const unresolvedSccs: CycleResolutionDebugScc[] = []

    for (const [sccIndex, scc] of cyclicSccs.entries()) {
      const sccFormulaVertices = scc.filter((vertex): vertex is FormulaVertex => vertex instanceof FormulaVertex)
      if (sccFormulaVertices.length === 0) {
        scc.forEach(vertex => {
          if (vertex instanceof RangeVertex) {
            vertex.clearCache()
          }
        })
        continue
      }

      scc.forEach(vertex => {
        if (vertex instanceof RangeVertex) {
          vertex.clearCache()
        }
      })

      // Probe dependencies for guard-like formulas with current values.
      const unresolvedAfterProbe = new Set<FormulaVertex>(sccFormulaVertices)
      const probeFailures = new Map<number, CycleResolutionDebugProbeFailure>()
      let madeProgress = true
      while (madeProgress && unresolvedAfterProbe.size > 0) {
        madeProgress = false
        for (const vertex of [...unresolvedAfterProbe]) {
          const probeAccessTracker = new ProbeAccessTracker()
          const probeValueResolver = this.createProbeValueResolver(probeAccessTracker)
          try {
            this.recomputeFormulaVertexValue(vertex, probeAccessTracker, probeValueResolver)
            unresolvedAfterProbe.delete(vertex)
            madeProgress = true
          } catch (e) {
            if (vertex.idInGraph !== undefined) {
              probeFailures.set(vertex.idInGraph, this.buildProbeFailureDebug(e, probeAccessTracker.snapshot()))
            }
            // The probe is best-effort. If dependency values are unavailable, retry after other vertices are processed.
          }
        }
      }

      const probeFailedFormulaIds = new Set<number>()
      unresolvedAfterProbe.forEach((vertex) => {
        if (vertex.idInGraph !== undefined) {
          probeFailedFormulaIds.add(vertex.idInGraph)
        }
      })

      const [acyclicOrder, unresolved] = this.resolveOrderFromActiveEdges(sccFormulaVertices, probeFailedFormulaIds)
      const finalUnresolved = new Set<FormulaVertex>(unresolved)

      acyclicOrder.forEach((vertex) => {
        try {
          if (changes !== undefined) {
            this.recomputeVertex(vertex, changes)
          } else {
            const newCellValue = this.recomputeFormulaVertexValue(vertex)
            const address = vertex.getAddress(this.lazilyTransformingAstService)
            this.columnSearch.add(getRawValue(newCellValue), address)
          }
          remainingCycledFormulas.delete(vertex)
          if (vertex.idInGraph !== undefined) {
            processedFormulaIds.add(vertex.idInGraph)
          }
        } catch (e) {
          remainingCycledFormulas.add(vertex)
          finalUnresolved.add(vertex)
          if (vertex.idInGraph !== undefined) {
            processedFormulaIds.add(vertex.idInGraph)
          }
        }
      })

      finalUnresolved.forEach((vertex) => {
        remainingCycledFormulas.add(vertex)
        if (vertex.idInGraph !== undefined) {
          processedFormulaIds.add(vertex.idInGraph)
        }
      })

      if (finalUnresolved.size > 0) {
        unresolvedSccs.push(this.buildCycleResolutionDebugScc(
          sccIndex,
          sccFormulaVertices,
          finalUnresolved,
          probeFailedFormulaIds,
          probeFailures,
        ))
      }
    }

    cycled.forEach((vertex: Vertex) => {
      if (!(vertex instanceof FormulaVertex)) {
        return
      }
      if (vertex.idInGraph !== undefined && !processedFormulaIds.has(vertex.idInGraph)) {
        remainingCycledFormulas.add(vertex)
      }
    })

    remainingCycledFormulas.forEach((vertex) => {
      if (changes !== undefined) {
        this.processVertexOnCycle(vertex, changes)
      } else {
        vertex.setCellValue(new CellError(ErrorType.CYCLE, undefined, vertex))
      }
    })

    this._cycleResolutionSnapshot = {
      totalCyclicSccCount: cyclicSccs.length,
      unresolvedSccCount: unresolvedSccs.length,
      unresolvedFormulaCount: unresolvedSccs.reduce((sum, scc) => sum + scc.unresolvedCount, 0),
      unresolvedSccs,
    }
  }

  private resolveOrderFromActiveEdges(sccFormulaVertices: FormulaVertex[], forceUnresolved: Set<number> = new Set<number>()): [FormulaVertex[], FormulaVertex[]] {
    const idToVertex = new Map<number, FormulaVertex>()
    sccFormulaVertices.forEach((vertex) => {
      if (vertex.idInGraph !== undefined) {
        idToVertex.set(vertex.idInGraph, vertex)
      }
    })

    const incoming = new Map<number, Set<number>>()
    const outgoing = new Map<number, Set<number>>()

    idToVertex.forEach((_, vertexId) => {
      incoming.set(vertexId, new Set())
      outgoing.set(vertexId, new Set())
    })

    const snapshot = this.activeEdgeCollector?.snapshot()
    if (snapshot !== undefined) {
      for (const [formulaId, deps] of snapshot.byFormula.entries()) {
        if (!idToVertex.has(formulaId)) {
          continue
        }
        const dependents = this.resolveDependenciesWithinScc(deps, idToVertex)
        for (const dependencyFormulaId of dependents) {
          if (dependencyFormulaId === formulaId) {
            incoming.get(formulaId)?.add(dependencyFormulaId)
            outgoing.get(dependencyFormulaId)?.add(formulaId)
            continue
          }
          incoming.get(formulaId)?.add(dependencyFormulaId)
          outgoing.get(dependencyFormulaId)?.add(formulaId)
        }
      }
    }

    const queue: number[] = []
    incoming.forEach((deps, nodeId) => {
      if (deps.size === 0 && !forceUnresolved.has(nodeId)) {
        queue.push(nodeId)
      }
    })

    const ordered: FormulaVertex[] = []
    while (queue.length > 0) {
      const node = queue.shift()!
      const vertex = idToVertex.get(node)
      if (vertex === undefined) {
        continue
      }
      ordered.push(vertex)
      const nextNodes = outgoing.get(node)
      if (nextNodes === undefined) {
        continue
      }
      nextNodes.forEach((nextNodeId) => {
        const nextIncoming = incoming.get(nextNodeId)
        if (nextIncoming === undefined) {
          return
        }
        nextIncoming.delete(node)
        if (nextIncoming.size === 0) {
          queue.push(nextNodeId)
        }
      })
    }

    const unresolved = sccFormulaVertices.filter((vertex) => {
      if (vertex.idInGraph === undefined) {
        return true
      }
      if (forceUnresolved.has(vertex.idInGraph)) {
        return true
      }
      return !ordered.includes(vertex)
    })

    return [ordered, unresolved]
  }

  private resolveDependenciesWithinScc(dependencies: ActiveDependency[], idToVertex: Map<number, FormulaVertex>): Set<number> {
    const dependencyIds = new Set<number>()
    const rangesWithPreciseDependencies = new Set<string>()

    for (const dependency of dependencies) {
      if (dependency.kind === 'RANGE_CELL') {
        rangesWithPreciseDependencies.add(this.rangeDependencyKey(dependency.start, dependency.end))
      }
    }

    for (const dependency of dependencies) {
      if (dependency.kind === 'CELL' || dependency.kind === 'NAMED_EXPRESSION' || dependency.kind === 'RANGE_CELL') {
        const vertex = this.dependencyGraph.getCell(dependency.address)
        if (vertex instanceof FormulaVertex && vertex.idInGraph !== undefined && idToVertex.has(vertex.idInGraph)) {
          dependencyIds.add(vertex.idInGraph)
        }
      } else if (dependency.kind === 'RANGE') {
        if (rangesWithPreciseDependencies.has(this.rangeDependencyKey(dependency.start, dependency.end))) {
          continue
        }
        for (const [vertexId, formulaVertex] of idToVertex.entries()) {
          const address = formulaVertex.getAddress(this.lazilyTransformingAstService)
          if (address.sheet !== dependency.start.sheet) {
            continue
          }
          if (address.col >= dependency.start.col && address.col <= dependency.end.col
            && address.row >= dependency.start.row && address.row <= dependency.end.row) {
            dependencyIds.add(vertexId)
          }
        }
      }
    }
    return dependencyIds
  }

  private rangeDependencyKey(start: SimpleCellAddress, end: SimpleCellAddress): string {
    return `${start.sheet}:${start.col}:${start.row}:${end.col}:${end.row}`
  }

  private buildCycleResolutionDebugScc(
    sccIndex: number,
    sccFormulaVertices: FormulaVertex[],
    finalUnresolved: Set<FormulaVertex>,
    probeFailedFormulaIds: Set<number>,
    probeFailures: Map<number, CycleResolutionDebugProbeFailure>,
  ): CycleResolutionDebugScc {
    const snapshot = this.activeEdgeCollector?.snapshot()
    const idToVertex = new Map<number, FormulaVertex>()
    sccFormulaVertices.forEach((vertex) => {
      if (vertex.idInGraph !== undefined) {
        idToVertex.set(vertex.idInGraph, vertex)
      }
    })

    const formulas = sccFormulaVertices.map((vertex) => this.buildCycleResolutionDebugFormula(
      vertex,
      finalUnresolved,
      probeFailedFormulaIds,
      idToVertex,
      snapshot,
      probeFailures,
    ))

    return {
      sccIndex,
      formulaCount: sccFormulaVertices.length,
      orderedCount: formulas.filter((formula) => formula.status === 'ordered').length,
      unresolvedCount: formulas.filter((formula) => formula.status === 'unresolved').length,
      unresolvedFormulaIds: formulas.filter((formula) => formula.status === 'unresolved' && formula.formulaId !== null).map((formula) => formula.formulaId as number),
      unresolvedAddresses: formulas.filter((formula) => formula.status === 'unresolved').map((formula) => formula.address),
      probeFailedFormulaIds: formulas.filter((formula) => formula.probeFailed && formula.formulaId !== null).map((formula) => formula.formulaId as number),
      dependencyCounts: {
        cell: formulas.reduce((sum, formula) => sum + formula.dependencyCounts.cell, 0),
        range: formulas.reduce((sum, formula) => sum + formula.dependencyCounts.range, 0),
        rangeCell: formulas.reduce((sum, formula) => sum + formula.dependencyCounts.rangeCell, 0),
        namedExpression: formulas.reduce((sum, formula) => sum + formula.dependencyCounts.namedExpression, 0),
        coarseRange: formulas.reduce((sum, formula) => sum + formula.dependencyCounts.coarseRange, 0),
      },
      reason: probeFailedFormulaIds.size > 0 ? 'probe_failed_or_still_cyclic' : 'still_cyclic_after_active_edges',
      formulas,
    }
  }

  private buildCycleResolutionDebugFormula(
    vertex: FormulaVertex,
    finalUnresolved: Set<FormulaVertex>,
    probeFailedFormulaIds: Set<number>,
    idToVertex: Map<number, FormulaVertex>,
    snapshot: ActiveEdgeSnapshot | undefined,
    probeFailures: Map<number, CycleResolutionDebugProbeFailure>,
  ): CycleResolutionDebugFormula {
    const address = vertex.getAddress(this.lazilyTransformingAstService)
    const formulaId = vertex.idInGraph ?? null
    const dependencies = formulaId !== null ? (snapshot?.byFormula.get(formulaId) ?? []) : []
    const preciseRangeKeys = this.collectPreciseRangeKeys(dependencies)
    const dependencyFormulaIds = formulaId !== null
      ? [...this.resolveDependenciesWithinScc(dependencies, idToVertex)].sort((a, b) => a - b)
      : []

    return {
      formulaId,
      sheet: this.dependencyGraph.sheetMapping.getSheetNameOrThrowError(address.sheet, { includePlaceholders: true }),
      address: this.formatAddress(address),
      formula: this.unparser.unparse(vertex.getFormula(this.lazilyTransformingAstService), address),
      status: finalUnresolved.has(vertex) ? 'unresolved' : 'ordered',
      probeFailed: formulaId !== null ? probeFailedFormulaIds.has(formulaId) : false,
      probeFailure: formulaId !== null ? probeFailures.get(formulaId) : undefined,
      dependencyCounts: {
        cell: dependencies.filter((dependency) => dependency.kind === 'CELL').length,
        range: dependencies.filter((dependency) => dependency.kind === 'RANGE').length,
        rangeCell: dependencies.filter((dependency) => dependency.kind === 'RANGE_CELL').length,
        namedExpression: dependencies.filter((dependency) => dependency.kind === 'NAMED_EXPRESSION').length,
        coarseRange: dependencies.filter((dependency) => dependency.kind === 'RANGE' && !preciseRangeKeys.has(this.rangeDependencyKey(dependency.start, dependency.end))).length,
      },
      sccDependencyFormulaIds: dependencyFormulaIds,
      sccDependencyAddresses: dependencyFormulaIds
        .map((dependencyId) => idToVertex.get(dependencyId))
        .filter((candidate): candidate is FormulaVertex => candidate !== undefined)
        .map((candidate) => this.formatAddress(candidate.getAddress(this.lazilyTransformingAstService))),
      activeDependencies: dependencies.map((dependency) => this.formatCycleDependencyDebug(dependency, preciseRangeKeys)),
    }
  }

  private formatCycleDependencyDebug(
    dependency: ActiveDependency,
    preciseRangeKeys: Set<string>,
  ): CycleResolutionDebugDependency {
    switch (dependency.kind) {
      case 'CELL':
        return {
          kind: dependency.kind,
          address: this.formatAddress(dependency.address),
        }
      case 'RANGE':
        return {
          kind: dependency.kind,
          range: this.formatRange(dependency.start, dependency.end),
          coveredByPreciseRange: preciseRangeKeys.has(this.rangeDependencyKey(dependency.start, dependency.end)),
        }
      case 'RANGE_CELL':
        return {
          kind: dependency.kind,
          range: this.formatRange(dependency.start, dependency.end),
          address: this.formatAddress(dependency.address),
        }
      case 'NAMED_EXPRESSION':
        return {
          kind: dependency.kind,
          expressionName: dependency.expressionName,
          address: this.formatAddress(dependency.address),
        }
    }
  }

  private collectPreciseRangeKeys(dependencies: ActiveDependency[]): Set<string> {
    const preciseRangeKeys = new Set<string>()
    for (const dependency of dependencies) {
      if (dependency.kind === 'RANGE_CELL') {
        preciseRangeKeys.add(this.rangeDependencyKey(dependency.start, dependency.end))
      }
    }
    return preciseRangeKeys
  }

  private formatAddress(address: SimpleCellAddress): string {
    return simpleCellAddressToString(
      this.dependencyGraph.sheetMapping.getSheetNameOrThrowError.bind(this.dependencyGraph.sheetMapping),
      address,
      -1,
    ) ?? `${address.sheet}:${address.col}:${address.row}`
  }

  private formatRange(start: SimpleCellAddress, end: SimpleCellAddress): string {
    return simpleCellRangeToString(
      this.dependencyGraph.sheetMapping.getSheetNameOrThrowError.bind(this.dependencyGraph.sheetMapping),
      {start, end},
      -1,
    ) ?? `${this.formatAddress(start)}:${this.formatAddress(end)}`
  }

  private static emptyCycleResolutionSnapshot(): CycleResolutionDebugSnapshot {
    return {
      totalCyclicSccCount: 0,
      unresolvedSccCount: 0,
      unresolvedFormulaCount: 0,
      unresolvedSccs: [],
    }
  }

  private buildProbeFailureDebug(error: unknown, lastAccess: ProbeAccess | undefined): CycleResolutionDebugProbeFailure {
    const errorMessage = error instanceof Error
      ? error.message
      : String(error)

    const errorType = error instanceof Error
      ? (error.name || error.constructor.name || 'Error')
      : typeof error

    const probeFailure: CycleResolutionDebugProbeFailure = {
      errorType,
      errorMessage,
      lastAccessKind: lastAccess?.kind ?? null,
    }

    if (lastAccess?.kind === 'CELL') {
      probeFailure.lastAccessAddress = this.formatAddress(lastAccess.address)
    } else if (lastAccess?.kind === 'RANGE') {
      probeFailure.lastAccessRange = this.formatRange(lastAccess.start, lastAccess.end)
    } else if (lastAccess?.kind === 'RANGE_CELL') {
      probeFailure.lastAccessAddress = this.formatAddress(lastAccess.address)
      probeFailure.lastAccessRange = this.formatRange(lastAccess.start, lastAccess.end)
    } else if (lastAccess?.kind === 'NAMED_EXPRESSION') {
      probeFailure.expressionName = lastAccess.expressionName
      probeFailure.lastAccessAddress = this.formatAddress(lastAccess.address)
    }

    return probeFailure
  }

  private createProbeValueResolver(probeAccessTracker?: ProbeAccessTracker): ProbeValueResolver {
    const verticesInProbe = new Set<FormulaVertex>()
    const resolver: ProbeValueResolver = {
      getCellValue: (address: SimpleCellAddress): InterpreterValue => {
        const vertex = this.dependencyGraph.getCell(address)
        if (vertex === undefined) {
          return EmptyValue
        }

        if (vertex instanceof FormulaVertex && !vertex.isComputed()) {
          if (verticesInProbe.has(vertex)) {
            throw Error(vertex instanceof ArrayFormulaVertex ? 'Array not computed yet.' : 'Value of the formula cell is not computed.')
          }

          verticesInProbe.add(vertex)
          try {
            this.recomputeFormulaVertexValue(vertex, probeAccessTracker, resolver)
          } finally {
            verticesInProbe.delete(vertex)
          }
        }

        if (vertex instanceof ArrayFormulaVertex) {
          return vertex.getArrayCellValue(address)
        }

        return vertex.getCellValue()
      },
    }

    return resolver
  }

  private recomputeFormulaVertexValue(vertex: FormulaVertex, probeAccessTracker?: ProbeAccessTracker, probeValueResolver?: ProbeValueResolver): InterpreterValue {
    const address = vertex.getAddress(this.lazilyTransformingAstService)
    if (vertex instanceof ArrayFormulaVertex && (vertex.array.size.isRef || !this.dependencyGraph.isThereSpaceForArray(vertex))) {
      return vertex.setNoSpace()
    } else {
      const formula = vertex.getFormula(this.lazilyTransformingAstService)
      const newCellValue = this.evaluateAstToCellValue(formula, new InterpreterState(address, this.config.useArrayArithmetic, vertex, this.activeEdgeCollector, probeAccessTracker, probeValueResolver))
      return vertex.setCellValue(newCellValue)
    }
  }

  private evaluateAstToCellValue(ast: Ast, state: InterpreterState): InterpreterValue {
    const interpreterValue = this.interpreter.evaluateAst(ast, state)
    if (interpreterValue instanceof SimpleRangeValue) {
      return interpreterValue
    } else if (interpreterValue === EmptyValue && this.config.evaluateNullToZero) {
      return 0
    } else {
      return interpreterValue
    }
  }
}
