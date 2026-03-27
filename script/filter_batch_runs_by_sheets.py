#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple


# Canonical sheet names + known mojibake aliases seen in this project.
CANONICAL_TO_ALIASES: Dict[str, Set[str]] = {
  "项目信息": {"项目信息", "椤圭洰淇℃伅"},
  "设计条件": {"设计条件", "璁捐鏉′欢"},
  "容器": {"容器", "瀹瑰櫒"},
  "功率计算": {"功率计算", "鍔熺巼璁＄畻"},
  "桨叶设计": {"桨叶设计", "妗ㄥ彾璁捐"},
  "轴设计": {"轴设计", "杞磋璁", "杞磋璁?"},
  "罐外单元": {"罐外单元", "缃愬鍗曞厓"},
  "驱动单元": {"驱动单元", "椹卞姩鍗曞厓"},
  "载荷": {"载荷", "杞借嵎"},
}

ALIAS_TO_CANONICAL: Dict[str, str] = {}
for canonical, aliases in CANONICAL_TO_ALIASES.items():
  for alias in aliases:
    ALIAS_TO_CANONICAL[alias] = canonical

ALL_ALIAS_TOKENS = sorted(ALIAS_TO_CANONICAL.keys(), key=len, reverse=True)


def normalize_sheet_name(name: str) -> str:
  text = (name or "").strip()
  if text in ALIAS_TO_CANONICAL:
    return ALIAS_TO_CANONICAL[text]
  for alias, canonical in ALIAS_TO_CANONICAL.items():
    if text.startswith(alias):
      return canonical
  return text


def extract_sheet_tokens(line: str) -> List[str]:
  """
  Extract one or more sheet tokens from a line.
  Handles accidental concatenation, e.g. "杞磋璁?缃愬鍗曞厓".
  """
  text = line.strip()
  if not text:
    return []
  found: List[str] = []
  for token in ALL_ALIAS_TOKENS:
    if token in text:
      found.append(token)
  if found:
    # Deduplicate while keeping deterministic order by token position.
    found = sorted(set(found), key=lambda t: text.find(t))
    return found
  return [text]


def load_target_sheets(path: Path) -> Tuple[Set[str], List[str]]:
  lines = path.read_text(encoding="utf-8-sig").splitlines()
  raw_tokens: List[str] = []
  for line in lines:
    raw_tokens.extend(extract_sheet_tokens(line))
  normalized = {normalize_sheet_name(token) for token in raw_tokens if token.strip()}
  return normalized, raw_tokens


def filter_one_run(payload: Dict[str, Any], target_sheets: Set[str]) -> Tuple[Dict[str, Any], int]:
  details = payload.get("mismatchDetails", []) or []
  filtered: List[Dict[str, Any]] = []
  for item in details:
    sheet = normalize_sheet_name(str(item.get("sheet", "")))
    if sheet in target_sheets:
      filtered.append(item)

  out = dict(payload)
  out["mismatchDetails"] = filtered
  out["filteredBySheets"] = sorted(target_sheets)
  out["originalMismatched"] = int(payload.get("mismatched", 0) or 0)
  out["mismatched"] = len(filtered)
  out["filteredOutMismatched"] = out["originalMismatched"] - out["mismatched"]
  return out, len(filtered)


def main() -> None:
  parser = argparse.ArgumentParser(description="Filter mismatchDetails in batch run JSON files by target sheet names.")
  parser.add_argument("--runs-dir", default="logs/batch_runs_100", help="Directory containing run_*.json files")
  parser.add_argument("--sheets-file", default="logs/sheets_9.txt", help="Text file with target sheet names")
  parser.add_argument("--out-dir", default="logs/batch_runs_100_filtered_9sheets", help="Output directory")
  args = parser.parse_args()

  runs_dir = Path(args.runs_dir).resolve()
  sheets_file = Path(args.sheets_file).resolve()
  out_dir = Path(args.out_dir).resolve()

  if not runs_dir.exists():
    raise SystemExit(f"runs dir not found: {runs_dir}")
  if not sheets_file.exists():
    raise SystemExit(f"sheets file not found: {sheets_file}")

  target_sheets, raw_tokens = load_target_sheets(sheets_file)
  out_dir.mkdir(parents=True, exist_ok=True)

  run_files = sorted(runs_dir.glob("run_*.json"))
  total_original = 0
  total_filtered = 0

  for run_file in run_files:
    payload = json.loads(run_file.read_text(encoding="utf-8-sig"))
    original = int(payload.get("mismatched", 0) or 0)
    filtered_payload, filtered_count = filter_one_run(payload, target_sheets)
    total_original += original
    total_filtered += filtered_count

    out_file = out_dir / run_file.name
    out_file.write_text(json.dumps(filtered_payload, ensure_ascii=False, indent=2), encoding="utf-8")

  summary = {
    "runsDir": str(runs_dir),
    "sheetsFile": str(sheets_file),
    "outDir": str(out_dir),
    "rawTokensFromFile": raw_tokens,
    "normalizedTargetSheets": sorted(target_sheets),
    "runFileCount": len(run_files),
    "totalOriginalMismatched": total_original,
    "totalFilteredMismatched": total_filtered,
  }
  (out_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

  print(f"Output dir: {out_dir}")
  print(f"Run files: {len(run_files)}")
  print(f"Original mismatched total: {total_original}")
  print(f"Filtered mismatched total: {total_filtered}")
  print(f"Summary: {out_dir / 'summary.json'}")


if __name__ == "__main__":
  main()
