#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import itertools
import json
import random
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple


@dataclass
class InputVar:
  key: str
  sheet: str
  address: str
  input_type: str
  baseline: Any
  candidates: List[Any]


def normalize_address(address: str) -> str:
  return address.replace("$", "").strip().upper()


def parse_number_like(raw: Any) -> Optional[float]:
  if raw is None:
    return None
  if isinstance(raw, bool):
    return None
  if isinstance(raw, (int, float)):
    return float(raw)
  text = str(raw).strip()
  if text == "":
    return None
  try:
    return float(text)
  except ValueError:
    return None


def dedup_keep_order(values: Sequence[Any]) -> List[Any]:
  out: List[Any] = []
  seen: Set[str] = set()
  for value in values:
    marker = json.dumps(value, ensure_ascii=False, sort_keys=True)
    if marker in seen:
      continue
    seen.add(marker)
    out.append(value)
  return out


def perturb_input_values(baseline: Any) -> List[Any]:
  num = parse_number_like(baseline)
  if num is not None:
    if float(num).is_integer():
      iv = int(num)
      return [iv, iv + 1, max(iv - 1, 0)]
    return [num, num * 1.1, num * 0.9]
  if isinstance(baseline, bool):
    return [baseline, not baseline]
  text = "" if baseline is None else str(baseline).strip()
  if text == "":
    return ["", "1"]
  return [text, f"{text}_test"]


def load_sheets_inputs(path: Path) -> Dict[str, InputVar]:
  payload = json.loads(path.read_text(encoding="utf-8-sig"))
  out: Dict[str, InputVar] = {}
  for sheet in payload.get("sheets", []):
    sheet_name = str(sheet.get("name", "")).strip()
    if sheet_name == "":
      continue
    for cell in sheet.get("cells", []):
      if cell.get("role") != "input":
        continue
      address = normalize_address(str(cell.get("address", "")))
      if address == "":
        continue
      key = f"{sheet_name}!{address}"
      baseline = cell.get("text")
      cands = dedup_keep_order(perturb_input_values(baseline))
      out[key] = InputVar(
        key=key,
        sheet=sheet_name,
        address=address,
        input_type="input",
        baseline=cands[0] if cands else "",
        candidates=cands if cands else [""],
      )
  return out


def load_dropdowns(path: Path) -> Dict[str, InputVar]:
  payload = json.loads(path.read_text(encoding="utf-8-sig"))
  out: Dict[str, InputVar] = {}
  for sheet in payload.get("sheets", []):
    sheet_name = str(sheet.get("sheetName", "")).strip()
    if sheet_name == "":
      continue
    for item in sheet.get("dropdowns", []):
      address = normalize_address(str(item.get("address", "")))
      if address == "":
        continue
      key = f"{sheet_name}!{address}"
      options: List[Any] = []
      for rule in item.get("rules", []):
        for opt in rule.get("options", []):
          opt_text = "" if opt is None else str(opt)
          options.append(opt_text)
      options = dedup_keep_order(options)
      if not options:
        options = [""]
      out[key] = InputVar(
        key=key,
        sheet=sheet_name,
        address=address,
        input_type="dropdown",
        baseline=options[0],
        candidates=options[:],
      )
  return out


def load_checkbox_links(path: Path) -> Dict[str, InputVar]:
  payload = json.loads(path.read_text(encoding="utf-8-sig"))
  out: Dict[str, InputVar] = {}
  for sheet in payload.get("sheets", []):
    sheet_name = str(sheet.get("sheetName", "")).strip()
    if sheet_name == "":
      continue
    for item in sheet.get("cells", []):
      address = normalize_address(str(item.get("linkedCell", "")))
      if address == "":
        continue
      key = f"{sheet_name}!{address}"
      possible = item.get("possibleValues", [])
      options = dedup_keep_order(possible if isinstance(possible, list) else [])
      if not options:
        options = [False, True]
      baseline = item.get("initialValue", options[0])
      options = dedup_keep_order([baseline] + options)
      out[key] = InputVar(
        key=key,
        sheet=sheet_name,
        address=address,
        input_type="checkbox_link",
        baseline=options[0],
        candidates=options,
      )
  return out


def merge_inputs(
  base_inputs: Dict[str, InputVar],
  dropdowns: Dict[str, InputVar],
  checks: Dict[str, InputVar],
) -> List[InputVar]:
  merged: Dict[str, InputVar] = dict(base_inputs)
  for key, var in dropdowns.items():
    merged[key] = var
  for key, var in checks.items():
    merged[key] = var
  return [merged[k] for k in sorted(merged.keys())]


def sample_inputs_diverse(
  inputs: Sequence[InputVar],
  max_inputs: int,
  seed: int,
) -> List[InputVar]:
  if max_inputs <= 0 or len(inputs) <= max_inputs:
    return list(inputs)

  rng = random.Random(seed)
  by_sheet: Dict[str, List[InputVar]] = defaultdict(list)
  by_type: Dict[str, List[InputVar]] = defaultdict(list)
  for var in inputs:
    by_sheet[var.sheet].append(var)
    by_type[var.input_type].append(var)

  for arr in by_sheet.values():
    rng.shuffle(arr)
  for arr in by_type.values():
    rng.shuffle(arr)

  selected: List[InputVar] = []
  selected_keys: Set[str] = set()

  def try_add(var: InputVar) -> bool:
    if var.key in selected_keys:
      return False
    if len(selected) >= max_inputs:
      return False
    selected.append(var)
    selected_keys.add(var.key)
    return True

  # 1) ensure every selected sheet has at least one sample
  for sheet in sorted(by_sheet.keys()):
    if len(selected) >= max_inputs:
      break
    try_add(by_sheet[sheet][0])

  # 2) ensure input type diversity (input/dropdown/checkbox_link)
  for input_type in sorted(by_type.keys()):
    if len(selected) >= max_inputs:
      break
    for var in by_type[input_type]:
      if try_add(var):
        break

  # 3) fill the rest with weighted random: larger sheets can contribute more
  pool: List[InputVar] = [v for v in inputs if v.key not in selected_keys]
  sheet_selected_count: Dict[str, int] = Counter(v.sheet for v in selected)
  while len(selected) < max_inputs and pool:
    weights: List[float] = []
    for v in pool:
      # Prefer broader coverage while allowing larger sheets to contribute more.
      sheet_size = len(by_sheet[v.sheet])
      penalty = 1.0 + float(sheet_selected_count.get(v.sheet, 0))
      weights.append(sheet_size / penalty)
    picked = rng.choices(pool, weights=weights, k=1)[0]
    try_add(picked)
    sheet_selected_count[picked.sheet] = sheet_selected_count.get(picked.sheet, 0) + 1
    pool = [v for v in pool if v.key not in selected_keys]

  return selected


def build_baseline_case(case_id: int) -> Dict[str, Any]:
  return {
    "id": case_id,
    "name": f"baseline_{case_id:04d}",
    "sets": [],
    "changedKeys": [],
  }


def build_smoke_cases(inputs: Sequence[InputVar], start_id: int) -> List[Dict[str, Any]]:
  cases: List[Dict[str, Any]] = [build_baseline_case(start_id)]
  case_id = start_id + 1
  for var in inputs:
    alt = next((v for v in var.candidates if v != var.baseline), None)
    if alt is None:
      continue
    cases.append({
      "id": case_id,
      "name": f"smoke_{case_id:04d}_{var.key}",
      "sets": [{"sheet": var.sheet, "address": var.address, "value": alt}],
      "changedKeys": [var.key],
    })
    case_id += 1
  return cases


def build_full_cases(inputs: Sequence[InputVar], start_id: int, max_runs: int) -> List[Dict[str, Any]]:
  cases: List[Dict[str, Any]] = [build_baseline_case(start_id)]
  case_id = start_id + 1
  domains = [var.candidates for var in inputs]
  for combo in itertools.product(*domains):
    sets = []
    changed = []
    for var, value in zip(inputs, combo):
      if value != var.baseline:
        sets.append({"sheet": var.sheet, "address": var.address, "value": value})
        changed.append(var.key)
    if not sets:
      continue
    cases.append({
      "id": case_id,
      "name": f"full_{case_id:04d}",
      "sets": sets,
      "changedKeys": changed,
    })
    case_id += 1
    if 0 < max_runs <= len(cases):
      break
  return cases


def pair_tuple(i: int, vi: Any, j: int, vj: Any) -> Tuple[int, str, int, str]:
  return (i, json.dumps(vi, ensure_ascii=False), j, json.dumps(vj, ensure_ascii=False))


def build_pairwise_cases(
  inputs: Sequence[InputVar],
  start_id: int,
  max_runs: int,
  seed: int,
  candidate_trials: int,
) -> List[Dict[str, Any]]:
  rng = random.Random(seed)
  cases: List[Dict[str, Any]] = [build_baseline_case(start_id)]
  case_id = start_id + 1
  n = len(inputs)
  uncovered: Set[Tuple[int, str, int, str]] = set()
  for i in range(n):
    for j in range(i + 1, n):
      for vi in inputs[i].candidates:
        for vj in inputs[j].candidates:
          uncovered.add(pair_tuple(i, vi, j, vj))

  if n == 0:
    return cases

  while uncovered:
    best_assign: List[Any] = [var.baseline for var in inputs]
    best_score = -1
    for _ in range(max(1, candidate_trials)):
      assign: List[Any] = []
      for var in inputs:
        assign.append(rng.choice(var.candidates))
      covered = 0
      for i in range(n):
        for j in range(i + 1, n):
          if pair_tuple(i, assign[i], j, assign[j]) in uncovered:
            covered += 1
      if covered > best_score:
        best_score = covered
        best_assign = assign
    if best_score <= 0:
      break

    sets = []
    changed = []
    for var, value in zip(inputs, best_assign):
      if value != var.baseline:
        sets.append({"sheet": var.sheet, "address": var.address, "value": value})
        changed.append(var.key)
    if sets:
      cases.append({
        "id": case_id,
        "name": f"pairwise_{case_id:04d}",
        "sets": sets,
        "changedKeys": changed,
      })
      case_id += 1

    for i in range(n):
      for j in range(i + 1, n):
        uncovered.discard(pair_tuple(i, best_assign[i], j, best_assign[j]))

    if 0 < max_runs <= len(cases):
      break
  return cases


def main() -> None:
  parser = argparse.ArgumentParser(description="Generate batch test cases for Excel vs HF compare.")
  parser.add_argument("--inputs", default="assets/sheets_inputs.json")
  parser.add_argument("--dropdowns", default="assets/dropdowns_resolved.json")
  parser.add_argument("--checks", default="assets/checkbox_linkcell_ranges.json")
  parser.add_argument("--output", default="logs/generated_cases.json")
  parser.add_argument("--mode", choices=["smoke", "pairwise", "full"], default="pairwise")
  parser.add_argument("--max-inputs", type=int, default=30, help="Limit variables used for case generation.")
  parser.add_argument("--max-candidates", type=int, default=3, help="Max values kept per input variable.")
  parser.add_argument("--max-runs", type=int, default=0, help="0 means unlimited.")
  parser.add_argument("--seed", type=int, default=42)
  parser.add_argument("--pairwise-trials", type=int, default=50)
  parser.add_argument("--sheet", action="append", default=[], help="Optional sheet filter (repeatable).")
  parser.add_argument("--sheet-file", default="", help="UTF-8 text file with one sheet name per line.")
  args = parser.parse_args()

  inputs_path = Path(args.inputs).resolve()
  dropdowns_path = Path(args.dropdowns).resolve()
  checks_path = Path(args.checks).resolve()
  output_path = Path(args.output).resolve()

  base_inputs = load_sheets_inputs(inputs_path)
  dropdowns = load_dropdowns(dropdowns_path)
  checks = load_checkbox_links(checks_path)

  merged = merge_inputs(base_inputs, dropdowns, checks)
  sheet_list = list(args.sheet)
  if args.sheet_file:
    sheet_file = Path(args.sheet_file).resolve()
    if not sheet_file.exists():
      raise SystemExit(f"Sheet file not found: {sheet_file}")
    for line in sheet_file.read_text(encoding="utf-8-sig").splitlines():
      name = line.strip()
      if name:
        sheet_list.append(name)
  if sheet_list:
    sheet_set = set(sheet_list)
    merged = [v for v in merged if v.sheet in sheet_set]

  for var in merged:
    var.candidates = dedup_keep_order(var.candidates)[:max(1, args.max_candidates)]
    if var.baseline not in var.candidates:
      var.candidates = dedup_keep_order([var.baseline] + var.candidates)

  if args.max_inputs > 0:
    merged = sample_inputs_diverse(merged, args.max_inputs, args.seed)

  if args.mode == "smoke":
    cases = build_smoke_cases(merged, 0)
  elif args.mode == "full":
    cases = build_full_cases(merged, 0, args.max_runs)
  else:
    cases = build_pairwise_cases(
      merged,
      start_id=0,
      max_runs=args.max_runs,
      seed=args.seed,
      candidate_trials=args.pairwise_trials,
    )

  if args.max_runs > 0:
    cases = cases[:args.max_runs]

  payload = {
    "meta": {
      "mode": args.mode,
      "seed": args.seed,
      "maxInputs": args.max_inputs,
      "maxCandidates": args.max_candidates,
      "inputCount": len(merged),
      "caseCount": len(cases),
    },
    "inputs": [
      {
        "key": v.key,
        "sheet": v.sheet,
        "address": v.address,
        "type": v.input_type,
        "baseline": v.baseline,
        "candidates": v.candidates,
      }
      for v in merged
    ],
    "cases": cases,
  }

  output_path.parent.mkdir(parents=True, exist_ok=True)
  output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
  print(f"Generated cases: {len(cases)}")
  print(f"Variables used: {len(merged)}")
  print(f"Output: {output_path}")


if __name__ == "__main__":
  main()
