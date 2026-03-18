#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compare real Excel-calculated values (xlwings) with HyperFormula outputs.

Usage example:
  python script/compare_excel_with_hf.py \
    --excel assets/table_test.xlsm \
    --hf-out logs/hf_eval.json \
    --report logs/excel_hf_compare_report.json
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sqlite3
import shutil
import subprocess
import tempfile
from datetime import date, datetime, time as dt_time
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
  import xlwings as xw
except Exception as exc:  # pragma: no cover
  raise SystemExit(
    "Missing dependency 'xlwings'. Install with: pip install xlwings"
  ) from exc

try:
  from openpyxl import load_workbook
except Exception as exc:  # pragma: no cover
  raise SystemExit(
    "Missing dependency 'openpyxl'. Install with: pip install openpyxl"
  ) from exc


def normalize(value: Any) -> Any:
  if value is None:
    return ""
  if isinstance(value, (datetime, date, dt_time)):
    return value.isoformat()
  if isinstance(value, float) and value.is_integer():
    return int(value)
  if isinstance(value, str):
    text = value.strip()
    if text == "":
      return ""
    return text
  return value


def values_equal(left: Any, right: Any, tolerance: float) -> bool:
  lhs = normalize(left)
  rhs = normalize(right)
  if isinstance(lhs, (int, float)) and isinstance(rhs, (int, float)):
    return math.isclose(float(lhs), float(rhs), rel_tol=0.0, abs_tol=tolerance)
  return lhs == rhs


def normalize_for_hf(value: Any) -> Any:
  if value is None:
    return ""
  if isinstance(value, bool):
    return value
  if isinstance(value, (int, float)):
    return value
  if isinstance(value, (datetime, date, dt_time)):
    return value.isoformat()
  text = str(value)
  return text


SET_EXPR_RE = re.compile(r"^(?P<sheet>.+)!(?P<addr>\$?[A-Za-z]{1,3}\$?\d+)$")
SHEET_REF_RE = re.compile(
  r"(?P<prefix>(^|[,(=+\-*/<>:&]))(?P<sheet>'(?:[^']|'')+'|[^'!,()=+\-*/<>:&]+(?: [^'!,()=+\-*/<>:&]+)*)!(?P<addr>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)",
  re.IGNORECASE,
)


def normalize_address(address: str) -> str:
  return address.replace("$", "").strip().upper()


def parse_a1_address(address: str) -> Tuple[int, int]:
  addr = normalize_address(address)
  idx = 0
  col = 0
  while idx < len(addr) and addr[idx].isalpha():
    col = col * 26 + (ord(addr[idx]) - ord("A") + 1)
    idx += 1
  if idx >= len(addr):
    raise ValueError(f"Invalid address: {address}")
  row = int(addr[idx:])
  return row, col


def parse_set_value(raw: str) -> Any:
  text = raw.strip()
  if text.lower() == "null":
    return None
  if text.lower() == "true":
    return True
  if text.lower() == "false":
    return False
  if (text.startswith("'") and text.endswith("'")) or (text.startswith('"') and text.endswith('"')):
    return text[1:-1]
  try:
    if "." in text:
      return float(text)
    return int(text)
  except ValueError:
    return text


def parse_set_expression(expr: str) -> Tuple[str, str, Any]:
  if "=" not in expr:
    raise ValueError(f"Invalid --set expression (missing '='): {expr}")
  lhs, rhs = expr.split("=", 1)
  lhs = lhs.strip()
  rhs = rhs.strip()
  m = SET_EXPR_RE.match(lhs)
  if not m:
    raise ValueError(f"Invalid --set target. Expected Sheet!A1, got: {lhs}")
  sheet = m.group("sheet").strip()
  address = normalize_address(m.group("addr"))
  value = parse_set_value(rhs)
  return sheet, address, value


def normalize_formula_sheet_refs(formula: str) -> str:
  def _repl(match: re.Match[str]) -> str:
    prefix = match.group("prefix")
    sheet = (match.group("sheet") or "").strip()
    addr = match.group("addr")
    if sheet.startswith("'") and sheet.endswith("'"):
      return f"{prefix}{sheet}!{addr}"
    escaped = sheet.replace("'", "''")
    return f"{prefix}'{escaped}'!{addr}"

  return SHEET_REF_RE.sub(_repl, formula)


def coerce_db_value(raw: Any) -> Any:
  if raw is None:
    return ""
  if isinstance(raw, bool):
    return raw
  if isinstance(raw, (int, float)):
    return raw
  text = str(raw).strip()
  if text == "":
    return ""
  lowered = text.lower()
  if lowered == "true":
    return True
  if lowered == "false":
    return False
  try:
    number = float(text)
    if number.is_integer():
      return int(number)
    return number
  except ValueError:
    return text


def apply_overrides_to_hf_input(payload: Dict[str, Any], overrides: List[Tuple[str, str, Any]]) -> None:
  if not overrides:
    return
  sheet_map = {sheet["name"]: sheet for sheet in payload.get("sheets", [])}
  formula_cell_set = {(c["sheet"], c["address"]) for c in payload.get("formulaCells", [])}

  for sheet_name, address, value in overrides:
    if sheet_name not in sheet_map:
      raise ValueError(f"--set sheet not found in workbook: {sheet_name}")

    row, col = parse_a1_address(address)
    matrix: List[List[Any]] = sheet_map[sheet_name]["cells"]
    while len(matrix) < row:
      matrix.append([])
    while len(matrix[row - 1]) < col:
      matrix[row - 1].append("")

    matrix[row - 1][col - 1] = normalize_for_hf(value)
    formula_cell_set.discard((sheet_name, address))

  payload["formulaCells"] = [
    cell for cell in payload.get("formulaCells", [])
    if (cell["sheet"], cell["address"]) in formula_cell_set
  ]


def col_to_letters(col: int) -> str:
  letters: List[str] = []
  n = col
  while n > 0:
    n, rem = divmod(n - 1, 26)
    letters.append(chr(65 + rem))
  return "".join(reversed(letters))


def to_a1(row: int, col: int) -> str:
  return f"{col_to_letters(col)}{row}"


def build_hf_input_from_workbook(excel_path: Path) -> Dict[str, Any]:
  wb = load_workbook(filename=str(excel_path), data_only=False, keep_vba=True, read_only=False)
  try:
    sheets: List[Dict[str, Any]] = []
    formula_cells: List[Dict[str, Any]] = []
    for ws in wb.worksheets:
      raw_cells: List[Tuple[int, int, Any]] = []
      max_row = 0
      max_col = 0

      # openpyxl max_row/max_column may include format-only tails in xlsm.
      # Scan concrete instantiated cells and keep only non-empty values.
      for cell in ws._cells.values():  # type: ignore[attr-defined]
        value = cell.value
        if value is None:
          continue
        if isinstance(value, str) and value == "":
          continue
        row = int(cell.row)
        col = int(cell.column)
        raw_cells.append((row, col, value))
        if row > max_row:
          max_row = row
        if col > max_col:
          max_col = col

      matrix: List[List[Any]] = []
      if max_row > 0 and max_col > 0:
        matrix = [["" for _ in range(max_col)] for _ in range(max_row)]
        for r, c, value in raw_cells:
          if isinstance(value, str) and value.startswith("="):
            matrix[r - 1][c - 1] = value
            formula_cells.append({
              "sheet": ws.title,
              "address": to_a1(r, c),
              "row": r - 1,
              "col": c - 1,
            })
          else:
            matrix[r - 1][c - 1] = normalize_for_hf(value)
      sheets.append({
        "name": ws.title,
        "cells": matrix,
      })
    return {
      "sheets": sheets,
      "formulaCells": formula_cells,
    }
  finally:
    wb.close()


def build_hf_input_from_db(db_path: Path, project_id: str | None) -> Dict[str, Any]:
  conn = sqlite3.connect(str(db_path))
  conn.row_factory = sqlite3.Row
  try:
    if project_id:
      rows = conn.execute(
        """
        SELECT c.sheet_name, c.cell_address, COALESCE(p.value, c.value) AS current_value, c.formula
        FROM cell_data c
        LEFT JOIN project_cell_values p
          ON p.project_id = ?
         AND p.sheet_name = c.sheet_name
         AND p.cell_address = c.cell_address
        ORDER BY c.sheet_name, c.cell_address
        """,
        (project_id,),
      ).fetchall()
    else:
      rows = conn.execute(
        """
        SELECT sheet_name, cell_address, value AS current_value, formula
        FROM cell_data
        ORDER BY sheet_name, cell_address
        """
      ).fetchall()

    by_sheet: Dict[str, List[sqlite3.Row]] = {}
    for row in rows:
      by_sheet.setdefault(str(row["sheet_name"]), []).append(row)

    sheets: List[Dict[str, Any]] = []
    formula_cells: List[Dict[str, Any]] = []

    for sheet_name, cells in by_sheet.items():
      max_row = 0
      max_col = 0
      coords: Dict[Tuple[int, int], Any] = {}
      for row in cells:
        address = normalize_address(str(row["cell_address"]))
        try:
          r, c = parse_a1_address(address)
        except Exception:
          continue
        if r <= 0 or c <= 0:
          continue
        max_row = max(max_row, r)
        max_col = max(max_col, c)

        formula = row["formula"]
        if formula is not None and str(formula).strip() != "":
          formula_text = str(formula).strip()
          if not formula_text.startswith("="):
            formula_text = f"={formula_text}"
          formula_text = normalize_formula_sheet_refs(formula_text)
          coords[(r, c)] = formula_text
          formula_cells.append({
            "sheet": sheet_name,
            "address": to_a1(r, c),
            "row": r - 1,
            "col": c - 1,
          })
        else:
          coords[(r, c)] = normalize_for_hf(coerce_db_value(row["current_value"]))

      matrix: List[List[Any]] = []
      if max_row > 0 and max_col > 0:
        matrix = [["" for _ in range(max_col)] for _ in range(max_row)]
        for (r, c), value in coords.items():
          matrix[r - 1][c - 1] = value
      sheets.append({
        "name": sheet_name,
        "cells": matrix,
      })

    return {
      "sheets": sheets,
      "formulaCells": formula_cells,
    }
  finally:
    conn.close()


def resolve_npm_command(npm_path: str | None) -> str:
  candidates = []
  if npm_path:
    candidates.append(npm_path)
  candidates.extend([
    shutil.which("npm"),
    shutil.which("npm.cmd"),
    r"C:\Program Files\nodejs\npm.cmd",
  ])

  for candidate in candidates:
    if candidate and Path(candidate).exists():
      return candidate

  raise FileNotFoundError(
    "Cannot find npm executable. Set PATH or pass --npm-path (e.g. C:\\Program Files\\nodejs\\npm.cmd)."
  )


def run_hf_eval(node_script: Path, hf_input: Path, hf_out: Path, npm_path: str | None) -> Dict[str, Any]:
  npm_cmd = resolve_npm_command(npm_path)
  cmd = [
    npm_cmd,
    "run",
    "tsnode",
    "--",
    str(node_script),
    str(hf_input),
    str(hf_out),
  ]
  completed = subprocess.run(cmd, check=False, capture_output=True, text=True)
  if completed.stdout.strip():
    print(completed.stdout.strip())
  if completed.stderr.strip():
    print(completed.stderr.strip())
  if completed.returncode != 0:
    raise RuntimeError(
      f"HF eval command failed with exit code {completed.returncode}\n"
      f"CMD: {' '.join(cmd)}\n"
      f"STDOUT:\n{completed.stdout}\n"
      f"STDERR:\n{completed.stderr}"
    )
  return json.loads(hf_out.read_text(encoding="utf-8"))


def read_excel_values(excel_path: Path, cells: List[Dict[str, Any]], overrides: List[Tuple[str, str, Any]]) -> Dict[Tuple[str, str], Any]:
  app = xw.App(visible=False, add_book=False)
  app.display_alerts = False
  app.screen_updating = False
  book = None
  try:
    book = app.books.open(str(excel_path), update_links=False, read_only=False)
    for sheet_name, address, value in overrides:
      book.sheets[sheet_name].range(address).value = value
    # Force real Excel recalculation.
    app.api.CalculateFullRebuild()

    output: Dict[Tuple[str, str], Any] = {}
    for item in cells:
      sheet_name = item["sheet"]
      address = item["address"]
      try:
        output[(sheet_name, address)] = book.sheets[sheet_name].range(address).value
      except Exception:
        output[(sheet_name, address)] = None
    return output
  finally:
    if book is not None:
      book.close()
    app.quit()


def main() -> None:
  parser = argparse.ArgumentParser(description="Compare Excel real calc vs HyperFormula outputs.")
  parser.add_argument("--excel", required=True, help="Path to xlsx/xlsm workbook.")
  parser.add_argument("--hf-source", choices=["db", "excel"], default="db", help="HF init data source.")
  parser.add_argument("--db-path", default="assets/app.db", help="SQLite DB path used when --hf-source=db.")
  parser.add_argument("--project-id", help="Optional project_id for project_cell_values overlay.")
  parser.add_argument("--node-script", default="script/hf_eval_sheets.ts")
  parser.add_argument("--npm-path", help="Optional explicit path to npm.cmd")
  parser.add_argument("--hf-out", default="logs/hf_eval.json")
  parser.add_argument("--report", default="logs/excel_hf_compare_report.json")
  parser.add_argument("--sheet", help="Optional sheet filter.")
  parser.add_argument("--set", action="append", default=[], help="Override one cell before compare, e.g. --set 罐外单元!AM77=1")
  parser.add_argument("--tolerance", type=float, default=1e-6)
  parser.add_argument("--max-mismatches", type=int, default=200)
  args = parser.parse_args()

  excel_path = Path(args.excel).resolve()
  db_path = Path(args.db_path).resolve()
  node_script = Path(args.node_script).resolve()
  hf_out = Path(args.hf_out).resolve()
  report_path = Path(args.report).resolve()

  if not excel_path.exists():
    raise SystemExit(f"Excel file not found: {excel_path}")
  if args.hf_source == "db" and not db_path.exists():
    raise SystemExit(f"DB file not found: {db_path}")
  if not node_script.exists():
    raise SystemExit(f"Node script not found: {node_script}")

  hf_out.parent.mkdir(parents=True, exist_ok=True)
  report_path.parent.mkdir(parents=True, exist_ok=True)
  overrides = [parse_set_expression(expr) for expr in args.set]

  if args.hf_source == "db":
    hf_input_payload = build_hf_input_from_db(db_path, args.project_id)
  else:
    hf_input_payload = build_hf_input_from_workbook(excel_path)
  apply_overrides_to_hf_input(hf_input_payload, overrides)
  with tempfile.NamedTemporaryFile(mode="w", suffix=".json", encoding="utf-8", delete=False) as tmp:
    hf_input_path = Path(tmp.name)
    json.dump(hf_input_payload, tmp, ensure_ascii=False)

  try:
    hf_payload = run_hf_eval(node_script, hf_input_path, hf_out, args.npm_path)
  finally:
    try:
      hf_input_path.unlink(missing_ok=True)
    except Exception:
      pass
  hf_cells = hf_payload.get("cells", [])
  if args.sheet:
    hf_cells = [cell for cell in hf_cells if cell.get("sheet") == args.sheet]

  excel_values = read_excel_values(excel_path, hf_cells, overrides)

  mismatch_details: List[Dict[str, Any]] = []
  matched = 0
  for item in hf_cells:
    key = (item["sheet"], item["address"])
    excel_value = excel_values.get(key)
    hf_value = item.get("hfValue")
    if values_equal(excel_value, hf_value, args.tolerance):
      matched += 1
      continue
    if len(mismatch_details) < args.max_mismatches:
      mismatch_details.append({
        "sheet": item["sheet"],
        "address": item["address"],
        "excel": normalize(excel_value),
        "hf": normalize(hf_value),
      })

  total = len(hf_cells)
  mismatches = total - matched
  report = {
    "excelPath": str(excel_path),
    "hfSource": args.hf_source,
    "dbPath": str(db_path) if args.hf_source == "db" else None,
    "projectId": args.project_id,
    "sheetFilter": args.sheet,
    "overrides": [{"sheet": s, "address": a, "value": normalize(v)} for (s, a, v) in overrides],
    "tolerance": args.tolerance,
    "totalFormulaCellsCompared": total,
    "matched": matched,
    "mismatched": mismatches,
    "matchRate": (matched / total) if total else 1.0,
    "mismatchDetails": mismatch_details,
  }
  report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

  print(f"Compared formula cells: {total}")
  print(f"Matched: {matched}")
  print(f"Mismatched: {mismatches}")
  print(f"Report: {report_path}")


if __name__ == "__main__":
  main()
