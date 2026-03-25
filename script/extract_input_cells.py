#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List


def extract_input_cells(payload: Dict[str, Any]) -> Dict[str, Any]:
  sheets = payload.get("sheets", [])
  if not isinstance(sheets, list):
    raise ValueError("Invalid JSON: `sheets` must be a list.")

  out_sheets: List[Dict[str, Any]] = []
  total_input = 0
  total_cells = 0

  for sheet in sheets:
    if not isinstance(sheet, dict):
      continue
    name = sheet.get("name", "")
    cells = sheet.get("cells", [])
    if not isinstance(cells, list):
      continue

    input_cells: List[Dict[str, Any]] = []
    for cell in cells:
      if not isinstance(cell, dict):
        continue
      total_cells += 1
      if cell.get("role") == "input":
        input_cells.append(cell)

    total_input += len(input_cells)
    out_sheets.append({
      "name": name,
      "cells": input_cells,
    })

  return {
    "meta": {
      "sheetCount": len(out_sheets),
      "totalCellsScanned": total_cells,
      "totalInputCells": total_input,
    },
    "sheets": out_sheets,
  }


def main() -> None:
  parser = argparse.ArgumentParser(description="Extract cells with role=input from sheets JSON.")
  parser.add_argument("--input", required=True, help="Path to source sheets JSON, e.g. assets/sheets_full.json")
  parser.add_argument("--output", required=True, help="Path to output JSON file")
  args = parser.parse_args()

  input_path = Path(args.input).resolve()
  output_path = Path(args.output).resolve()

  if not input_path.exists():
    raise SystemExit(f"Input file not found: {input_path}")

  payload = json.loads(input_path.read_text(encoding="utf-8-sig"))
  result = extract_input_cells(payload)

  output_path.parent.mkdir(parents=True, exist_ok=True)
  output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

  print(f"Input: {input_path}")
  print(f"Output: {output_path}")
  print(f"Sheets: {result['meta']['sheetCount']}")
  print(f"Total cells scanned: {result['meta']['totalCellsScanned']}")
  print(f"Input cells extracted: {result['meta']['totalInputCells']}")


if __name__ == "__main__":
  main()
