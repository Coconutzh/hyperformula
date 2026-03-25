#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import math
import sqlite3
import subprocess
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


def as_number(value: Any) -> Optional[float]:
  if isinstance(value, bool):
    return None
  if isinstance(value, (int, float)):
    return float(value)
  try:
    return float(str(value))
  except Exception:
    return None


def classify_reason(excel: Any, hf: Any, tolerance: float) -> str:
  if isinstance(hf, dict) and hf.get("errorType"):
    return f"hf_error_{hf.get('errorType')}"
  if isinstance(excel, str) and excel.startswith("#"):
    return "excel_error"
  en = as_number(excel)
  hn = as_number(hf)
  if en is not None and hn is not None:
    if not math.isclose(en, hn, rel_tol=0.0, abs_tol=tolerance):
      return "numeric_diff"
    return "same_numeric"
  if str(excel) != str(hf):
    return "text_or_type_mismatch"
  return "unknown"


def load_formula_map(db_path: Path) -> Dict[str, str]:
  conn = sqlite3.connect(str(db_path))
  try:
    rows = conn.execute("SELECT sheet_name, cell_address, formula FROM cell_data WHERE formula IS NOT NULL AND TRIM(formula) <> ''").fetchall()
    out: Dict[str, str] = {}
    for sheet, addr, formula in rows:
      key = f"{sheet}!{str(addr).replace('$', '').upper()}"
      out[key] = str(formula)
    return out
  finally:
    conn.close()


def run_one_case(
  compare_script: Path,
  excel_path: Path,
  out_report: Path,
  case_sets: List[Dict[str, Any]],
  tolerance: float,
  npm_path: Optional[str],
  hf_source: str,
  db_path: Path,
  project_id: Optional[str],
) -> Dict[str, Any]:
  cmd: List[str] = [
    "python",
    str(compare_script),
    "--excel",
    str(excel_path),
    "--report",
    str(out_report),
    "--tolerance",
    str(tolerance),
    "--hf-source",
    hf_source,
  ]
  if db_path:
    cmd.extend(["--db-path", str(db_path)])
  if project_id:
    cmd.extend(["--project-id", project_id])
  if npm_path:
    cmd.extend(["--npm-path", npm_path])
  for item in case_sets:
    cmd.extend(["--set", f"{item['sheet']}!{item['address']}={item['value']}"])

  completed = subprocess.run(cmd, check=False, capture_output=True, text=True)
  if completed.returncode != 0:
    raise RuntimeError(
      f"Case command failed with exit code {completed.returncode}\n"
      f"CMD: {' '.join(cmd)}\n"
      f"STDOUT:\n{completed.stdout}\n"
      f"STDERR:\n{completed.stderr}"
    )
  return json.loads(out_report.read_text(encoding="utf-8"))


def main() -> None:
  parser = argparse.ArgumentParser(description="Run batch Excel vs HF compare and aggregate deduped errors.")
  parser.add_argument("--cases", default="logs/generated_cases.json")
  parser.add_argument("--excel", default="assets/table_test.xlsm")
  parser.add_argument("--compare-script", default="script/compare_excel_with_hf.py")
  parser.add_argument("--runs-dir", default="logs/batch_runs")
  parser.add_argument("--summary-out", default="logs/batch_summary.json")
  parser.add_argument("--errors-out", default="logs/batch_errors_dedup.json")
  parser.add_argument("--db-path", default="assets/app.db")
  parser.add_argument("--project-id")
  parser.add_argument("--hf-source", choices=["db", "excel"], default="db")
  parser.add_argument("--npm-path")
  parser.add_argument("--tolerance", type=float, default=1e-6)
  parser.add_argument("--max-runs", type=int, default=0, help="0 means run all cases")
  args = parser.parse_args()

  cases_path = Path(args.cases).resolve()
  excel_path = Path(args.excel).resolve()
  compare_script = Path(args.compare_script).resolve()
  runs_dir = Path(args.runs_dir).resolve()
  summary_out = Path(args.summary_out).resolve()
  errors_out = Path(args.errors_out).resolve()
  db_path = Path(args.db_path).resolve()

  if not cases_path.exists():
    raise SystemExit(f"Cases file not found: {cases_path}")
  if not excel_path.exists():
    raise SystemExit(f"Excel file not found: {excel_path}")
  if not compare_script.exists():
    raise SystemExit(f"Compare script not found: {compare_script}")
  if args.hf_source == "db" and not db_path.exists():
    raise SystemExit(f"DB file not found: {db_path}")

  payload = json.loads(cases_path.read_text(encoding="utf-8-sig"))
  cases: List[Dict[str, Any]] = payload.get("cases", [])
  if args.max_runs > 0:
    cases = cases[:args.max_runs]

  runs_dir.mkdir(parents=True, exist_ok=True)
  summary_out.parent.mkdir(parents=True, exist_ok=True)
  errors_out.parent.mkdir(parents=True, exist_ok=True)

  formula_map = load_formula_map(db_path) if db_path.exists() else {}
  dedup: Dict[str, Dict[str, Any]] = {}

  total_compared = 0
  total_matched = 0
  total_mismatched = 0
  run_results: List[Dict[str, Any]] = []

  for idx, case in enumerate(cases, start=1):
    run_report = runs_dir / f"run_{idx:04d}.json"
    report = run_one_case(
      compare_script=compare_script,
      excel_path=excel_path,
      out_report=run_report,
      case_sets=case.get("sets", []),
      tolerance=args.tolerance,
      npm_path=args.npm_path,
      hf_source=args.hf_source,
      db_path=db_path,
      project_id=args.project_id,
    )

    compared = int(report.get("totalFormulaCellsCompared", 0))
    matched = int(report.get("matched", 0))
    mismatched = int(report.get("mismatched", 0))
    total_compared += compared
    total_matched += matched
    total_mismatched += mismatched

    run_results.append({
      "runIndex": idx,
      "caseId": case.get("id"),
      "name": case.get("name"),
      "changedKeys": case.get("changedKeys", []),
      "compared": compared,
      "matched": matched,
      "mismatched": mismatched,
      "reportPath": str(run_report),
    })

    for item in report.get("mismatchDetails", []):
      key = f"{item.get('sheet')}!{item.get('address')}"
      reason = classify_reason(item.get("excel"), item.get("hf"), args.tolerance)
      rec = dedup.get(key)
      if rec is None:
        rec = {
          "sheet": item.get("sheet"),
          "address": item.get("address"),
          "formula": formula_map.get(key),
          "errorCount": 0,
          "reasonCounter": Counter(),
          "examples": [],
          "firstSeenRun": idx,
          "lastSeenRun": idx,
        }
        dedup[key] = rec
      rec["errorCount"] += 1
      rec["lastSeenRun"] = idx
      rec["reasonCounter"][reason] += 1
      if len(rec["examples"]) < 5:
        rec["examples"].append({
          "runIndex": idx,
          "excel": item.get("excel"),
          "hf": item.get("hf"),
          "reason": reason,
        })

    print(f"[{idx}/{len(cases)}] mismatched={mismatched}, matched={matched}, compared={compared}")

  dedup_rows: List[Dict[str, Any]] = []
  for _, rec in sorted(dedup.items(), key=lambda kv: kv[1]["errorCount"], reverse=True):
    reason_counter: Counter = rec["reasonCounter"]
    top_reasons = [{"reason": r, "count": c} for r, c in reason_counter.most_common(5)]
    dedup_rows.append({
      "sheet": rec["sheet"],
      "address": rec["address"],
      "formula": rec["formula"],
      "errorCount": rec["errorCount"],
      "firstSeenRun": rec["firstSeenRun"],
      "lastSeenRun": rec["lastSeenRun"],
      "topReasons": top_reasons,
      "examples": rec["examples"],
    })

  summary = {
    "casesFile": str(cases_path),
    "excelPath": str(excel_path),
    "hfSource": args.hf_source,
    "dbPath": str(db_path) if args.hf_source == "db" else None,
    "projectId": args.project_id,
    "runCount": len(cases),
    "totalCompared": total_compared,
    "totalMatched": total_matched,
    "totalMismatched": total_mismatched,
    "matchRate": (total_matched / total_compared) if total_compared else 1.0,
    "uniqueErrorCells": len(dedup_rows),
    "runs": run_results,
  }

  summary_out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
  errors_out.write_text(json.dumps({"rows": dedup_rows}, ensure_ascii=False, indent=2), encoding="utf-8")

  print(f"Summary: {summary_out}")
  print(f"Dedup errors: {errors_out}")


if __name__ == "__main__":
  main()
