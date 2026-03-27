#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List


def to_key(value: Any) -> str:
  return json.dumps(value, ensure_ascii=False, sort_keys=True)


def normalize_sheet(sheet: str) -> str:
  text = (sheet or "").strip()
  alias = {
    "妗ㄥ彾璁捐": "桨叶设计",
    "缃愬鍗曞厓": "罐外单元",
    "椹卞姩鍗曞厓": "驱动单元",
    "鍔熺巼璁＄畻": "功率计算",
    "璁捐鏉′欢": "设计条件",
    "瀹瑰櫒": "容器",
    "椤圭洰淇℃伅": "项目信息",
    "杞借嵎": "载荷",
    "杞磋璁": "轴设计",
    "杞磋璁?": "轴设计",
  }
  if text in alias:
    return alias[text]
  for k, v in alias.items():
    if text.startswith(k):
      return v
  return text


def main() -> None:
  parser = argparse.ArgumentParser(description="Aggregate filtered mismatchDetails by unique cell.")
  parser.add_argument("--runs-dir", default="logs/batch_runs_100_filtered_9sheets")
  parser.add_argument("--out-json", default="logs/batch_runs_100_filtered_9sheets_errors_summary.json")
  args = parser.parse_args()

  runs_dir = Path(args.runs_dir).resolve()
  out_json = Path(args.out_json).resolve()
  if not runs_dir.exists():
    raise SystemExit(f"runs dir not found: {runs_dir}")

  agg: Dict[str, Dict[str, Any]] = {}
  total_events = 0
  run_files = sorted(runs_dir.glob("run_*.json"))

  for run_file in run_files:
    run_payload = json.loads(run_file.read_text(encoding="utf-8-sig"))
    run_index_str = run_file.stem.replace("run_", "")
    try:
      run_index = int(run_index_str)
    except ValueError:
      run_index = None

    for item in run_payload.get("mismatchDetails", []) or []:
      total_events += 1
      sheet = normalize_sheet(str(item.get("sheet", "")))
      address = str(item.get("address", ""))
      cell_key = f"{sheet}!{address}"
      excel_value = item.get("excel")
      hf_value = item.get("hf")

      rec = agg.get(cell_key)
      if rec is None:
        rec = {
          "sheet": sheet,
          "address": address,
          "mismatchCount": 0,
          "runIndices": [],
          "excelValueMap": {},
          "hfValueMap": {},
          "hfErrorTypeMap": {},
        }
        agg[cell_key] = rec

      rec["mismatchCount"] += 1
      if run_index is not None:
        rec["runIndices"].append(run_index)

      ek = to_key(excel_value)
      if ek not in rec["excelValueMap"]:
        rec["excelValueMap"][ek] = excel_value

      hk = to_key(hf_value)
      if hk not in rec["hfValueMap"]:
        rec["hfValueMap"][hk] = hf_value

      if isinstance(hf_value, dict) and hf_value.get("errorType"):
        et = str(hf_value.get("errorType"))
        rec["hfErrorTypeMap"][et] = et

  rows: List[Dict[str, Any]] = []
  for _, rec in sorted(agg.items(), key=lambda kv: kv[1]["mismatchCount"], reverse=True):
    run_indices = sorted(set(rec["runIndices"]))
    rows.append({
      "sheet": rec["sheet"],
      "address": rec["address"],
      "mismatchCount": rec["mismatchCount"],
      "firstSeenRun": run_indices[0] if run_indices else None,
      "lastSeenRun": run_indices[-1] if run_indices else None,
      "excelValues": list(rec["excelValueMap"].values()),
      "hfValues": list(rec["hfValueMap"].values()),
      "hfErrorTypes": sorted(rec["hfErrorTypeMap"].keys()),
    })

  out_payload = {
    "runsDir": str(runs_dir),
    "runFileCount": len(run_files),
    "totalMismatchEvents": total_events,
    "uniqueMismatchCells": len(rows),
    "rows": rows,
  }
  out_json.parent.mkdir(parents=True, exist_ok=True)
  out_json.write_text(json.dumps(out_payload, ensure_ascii=False, indent=2), encoding="utf-8")

  print(f"Output: {out_json}")
  print(f"Run files: {len(run_files)}")
  print(f"Total mismatch events: {total_events}")
  print(f"Unique mismatch cells: {len(rows)}")


if __name__ == "__main__":
  main()
