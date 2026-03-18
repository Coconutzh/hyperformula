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
import {InterpreterState} from './interpreter/InterpreterState'
import {EmptyValue, getRawValue, InterpreterValue} from './interpreter/InterpreterValue'
import {SimpleRangeValue} from './SimpleRangeValue'
import {LazilyTransformingAstService} from './LazilyTransformingAstService'
import {ColumnSearchStrategy} from './Lookup/SearchStrategy'
import {Ast, RelativeDependency} from './parser'
import {Statistics, StatType} from './statistics'

export class Evaluator {
  private activeEdgeCollector?: ActiveEdgeCollector
  private _activeEdgeSnapshot?: ActiveEdgeSnapshot

  constructor(
    private readonly config: Config,
    private readonly stats: Statistics,
    public readonly interpreter: Interpreter,
    private readonly lazilyTransformingAstService: LazilyTransformingAstService,
    private readonly dependencyGraph: DependencyGraph,
    private readonly columnSearch: ColumnSearchStrategy,
  ) {
  }

  public get activeEdgeSnapshot(): ActiveEdgeSnapshot | undefined {
    return this._activeEdgeSnapshot
  }

  public run(): void {
    this.activeEdgeCollector = new ActiveEdgeCollector()
    this.stats.start(StatType.TOP_SORT)
    const {sorted, cycled, cyclicSccs} = this.dependencyGraph.topSortWithScc()
    this.stats.end(StatType.TOP_SORT)

    this.stats.measure(StatType.EVALUATION, () => {
      this.recomputeFormulas(cycled, sorted, cyclicSccs)
    })
    this._activeEdgeSnapshot = this.activeEdgeCollector.snapshot()
    this.activeEdgeCollector = undefined
  }

  public partialRun(vertices: Vertex[]): ContentChanges {
    this.activeEdgeCollector = new ActiveEdgeCollector()
    const changes = ContentChanges.empty()

    this.stats.measure(StatType.EVALUATION, () => {
      this.dependencyGraph.graph.getTopSortedWithSccSubgraphFrom(vertices,
        (vertex: Vertex) => this.recomputeVertex(vertex, changes),
        (vertex: Vertex) => this.processVertexOnCycle(vertex, changes),
      )
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
  private recomputeFormulas(cycled: Vertex[], sorted: Vertex[], cyclicSccs: Vertex[][]): void {
    sorted.forEach((vertex: Vertex) => {
      if (vertex instanceof FormulaVertex) {
        const newCellValue = this.recomputeFormulaVertexValue(vertex)
        const address = vertex.getAddress(this.lazilyTransformingAstService)
        this.columnSearch.add(getRawValue(newCellValue), address)
      } else if (vertex instanceof RangeVertex) {
        vertex.clearCache()
      }
    })

    this.resolveCyclicSccs(cycled, cyclicSccs)
  }

  private resolveCyclicSccs(cycled: Vertex[], cyclicSccs: Vertex[][]): void {
    const remainingCycledFormulas = new Set<FormulaVertex>(cycled.filter((vertex): vertex is FormulaVertex => vertex instanceof FormulaVertex))
    const processedFormulaIds = new Set<number>()

    for (const scc of cyclicSccs) {
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
      sccFormulaVertices.forEach((vertex) => {
        if (!vertex.isComputed()) {
          try {
            this.recomputeFormulaVertexValue(vertex)
          } catch (e) {
            // The probe is best-effort. If dependency values are unavailable, keep conservative cycle handling.
          }
        }
      })

      const [acyclicOrder, unresolved] = this.resolveOrderFromActiveEdges(sccFormulaVertices)

      acyclicOrder.forEach((vertex) => {
        const newCellValue = this.recomputeFormulaVertexValue(vertex)
        const address = vertex.getAddress(this.lazilyTransformingAstService)
        this.columnSearch.add(getRawValue(newCellValue), address)
        remainingCycledFormulas.delete(vertex)
        if (vertex.idInGraph !== undefined) {
          processedFormulaIds.add(vertex.idInGraph)
        }
      })

      unresolved.forEach((vertex) => {
        remainingCycledFormulas.add(vertex)
        if (vertex.idInGraph !== undefined) {
          processedFormulaIds.add(vertex.idInGraph)
        }
      })
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
      vertex.setCellValue(new CellError(ErrorType.CYCLE, undefined, vertex))
    })
  }

  private resolveOrderFromActiveEdges(sccFormulaVertices: FormulaVertex[]): [FormulaVertex[], FormulaVertex[]] {
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
      if (deps.size === 0) {
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
      return !ordered.includes(vertex)
    })

    return [ordered, unresolved]
  }

  private resolveDependenciesWithinScc(dependencies: ActiveDependency[], idToVertex: Map<number, FormulaVertex>): Set<number> {
    const dependencyIds = new Set<number>()
    for (const dependency of dependencies) {
      if (dependency.kind === 'CELL' || dependency.kind === 'NAMED_EXPRESSION') {
        const vertex = this.dependencyGraph.getCell(dependency.address)
        if (vertex instanceof FormulaVertex && vertex.idInGraph !== undefined && idToVertex.has(vertex.idInGraph)) {
          dependencyIds.add(vertex.idInGraph)
        }
      } else if (dependency.kind === 'RANGE') {
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

  private recomputeFormulaVertexValue(vertex: FormulaVertex): InterpreterValue {
    const address = vertex.getAddress(this.lazilyTransformingAstService)
    if (vertex instanceof ArrayFormulaVertex && (vertex.array.size.isRef || !this.dependencyGraph.isThereSpaceForArray(vertex))) {
      return vertex.setNoSpace()
    } else {
      const formula = vertex.getFormula(this.lazilyTransformingAstService)
      const newCellValue = this.evaluateAstToCellValue(formula, new InterpreterState(address, this.config.useArrayArithmetic, vertex, this.activeEdgeCollector))
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
