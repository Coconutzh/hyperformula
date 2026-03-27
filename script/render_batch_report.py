#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import html
import json
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Tuple


def bar(pct: float, width: int = 24) -> str:
  pct = max(0.0, min(1.0, pct))
  n = round(pct * width)
  return "█" * n + "░" * (width - n)


def pct_text(v: float) -> str:
  return f"{v:.2%}"


def sheet_alias(raw: str) -> str:
  # Summary changedKeys now carry proper Chinese sheet names.
  # Keep a tiny fallback map for known mojibake variants from historical files.
  fallback = {
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
  val = (raw or "").strip()
  if val in fallback:
    return fallback[val]
  for k, v in fallback.items():
    if val.startswith(k):
      return v
  return val


def build_stats(summary: Dict[str, Any], dedup: Dict[str, Any]) -> Dict[str, Any]:
  runs = summary.get("runs", [])
  rows = dedup.get("rows", [])

  change_event_total = 0
  unique_keys = set()
  sheet_event = Counter()
  sheet_unique: Dict[str, set] = defaultdict(set)

  for run in runs:
    keys = run.get("changedKeys", []) or []
    change_event_total += len(keys)
    for key in keys:
      text = str(key)
      if "!" in text:
        sheet_raw, addr = text.split("!", 1)
        sheet = sheet_alias(sheet_raw)
        uniq = f"{sheet}!{addr}"
      else:
        sheet = sheet_alias(text)
        uniq = text
      unique_keys.add(uniq)
      sheet_event[sheet] += 1
      sheet_unique[sheet].add(uniq)

  run_count = int(summary.get("runCount", 0) or 0)
  compared = int(summary.get("totalCompared", 0) or 0)
  matched = int(summary.get("totalMatched", 0) or 0)
  mismatched = int(summary.get("totalMismatched", 0) or 0)
  match_rate = float(summary.get("matchRate", 0.0) or 0.0) if compared else 0.0
  mismatch_rate = 1.0 - match_rate if compared else 0.0

  dedup_rows = len(rows)
  repeated = 0
  single = 0
  buckets = {"1次": 0, "2-5次": 0, "6-20次": 0, "21-50次": 0, "51次以上": 0}
  reasons = Counter()

  for row in rows:
    ec = int(row.get("errorCount", 0) or 0)
    if ec <= 1:
      single += 1
    else:
      repeated += 1
    if ec == 1:
      buckets["1次"] += 1
    elif ec <= 5:
      buckets["2-5次"] += 1
    elif ec <= 20:
      buckets["6-20次"] += 1
    elif ec <= 50:
      buckets["21-50次"] += 1
    else:
      buckets["51次以上"] += 1

    for item in row.get("topReasons", []) or []:
      reasons[str(item.get("reason", "unknown"))] += int(item.get("count", 0) or 0)

  return {
    "runs": runs,
    "rows": rows,
    "run_count": run_count,
    "compared": compared,
    "matched": matched,
    "mismatched": mismatched,
    "match_rate": match_rate,
    "mismatch_rate": mismatch_rate,
    "change_event_total": change_event_total,
    "unique_change_cells": len(unique_keys),
    "covered_sheets": len(sheet_event),
    "sheet_event": sheet_event,
    "sheet_unique": sheet_unique,
    "dedup_rows": dedup_rows,
    "repeated": repeated,
    "single": single,
    "repeat_ratio": (repeated / dedup_rows) if dedup_rows else 0.0,
    "single_ratio": (single / dedup_rows) if dedup_rows else 0.0,
    "buckets": buckets,
    "reasons": reasons,
  }


def render_markdown(stats: Dict[str, Any]) -> str:
  reason_total = sum(stats["reasons"].values())
  top_repeat = sorted(stats["rows"], key=lambda x: int(x.get("errorCount", 0) or 0), reverse=True)[:15]

  lines: List[str] = []
  lines += [
    "# Hyperformula 批量测试汇报（100批）",
    "",
    "## 一眼看懂",
    "",
    f"- 本次共执行 **{stats['run_count']}** 批测试，比较公式单元格 **{stats['compared']:,}** 个。",
    f"- 结果：匹配 **{stats['matched']:,}**，不匹配 **{stats['mismatched']:,}**，匹配率 **{pct_text(stats['match_rate'])}**。",
    f"- 变更覆盖：共发生 **{stats['change_event_total']:,}** 次改单元格事件，涉及 **{stats['unique_change_cells']:,}** 个唯一改单元格，覆盖 **{stats['covered_sheets']}** 个表。",
    f"- 去重后 mismatch 单元格 **{stats['dedup_rows']}** 个，其中重复出现（>1次）**{stats['repeated']}** 个，占比 **{pct_text(stats['repeat_ratio'])}**。",
    "",
    "## （1）修改单元格范围与总体结果",
    "",
    "### 1) 修改覆盖总览",
    "",
    "| 指标 | 数值 |",
    "|---|---:|",
    f"| 测试批次（runCount） | {stats['run_count']} |",
    f"| 修改事件总数（changedKeys条目总和） | {stats['change_event_total']:,} |",
    f"| 唯一修改单元格数（去重后） | {stats['unique_change_cells']:,} |",
    f"| 覆盖表格数 | {stats['covered_sheets']} |",
    "",
    "### 2) 各表修改占比（按事件 + 按唯一改单元格）",
    "",
    "| 表格 | 修改事件数 | 事件占比 | 唯一改单元格数 | 唯一占比 | 事件占比可视化 |",
    "|---|---:|---:|---:|---:|---|",
  ]
  for sheet, cnt in stats["sheet_event"].most_common():
    e_pct = (cnt / stats["change_event_total"]) if stats["change_event_total"] else 0.0
    u_cnt = len(stats["sheet_unique"][sheet])
    u_pct = (u_cnt / stats["unique_change_cells"]) if stats["unique_change_cells"] else 0.0
    lines.append(f"| {sheet} | {cnt:,} | {pct_text(e_pct)} | {u_cnt} | {pct_text(u_pct)} | `{bar(e_pct)}` |")

  lines += [
    "",
    "### 3) Compared / Matched / Mismatched",
    "",
    "| 指标 | 数值 | 占比 | 可视化 |",
    "|---|---:|---:|---|",
    f"| Compared | {stats['compared']:,} | 100.00% | `{bar(1.0)}` |",
    f"| Matched | {stats['matched']:,} | {pct_text(stats['match_rate'])} | `{bar(stats['match_rate'])}` |",
    f"| Mismatched | {stats['mismatched']:,} | {pct_text(stats['mismatch_rate'])} | `{bar(stats['mismatch_rate'])}` |",
    "",
    "## （2）Mismatch 去重重复情况与原因汇总",
    "",
    "### 1) 重复出现情况",
    "",
    "| 指标 | 数值 | 占比 |",
    "|---|---:|---:|",
    f"| 去重后 mismatch 单元格总数 | {stats['dedup_rows']} | 100.00% |",
    f"| 重复出现单元格（errorCount > 1） | {stats['repeated']} | {pct_text(stats['repeat_ratio'])} |",
    f"| 仅出现一次单元格（errorCount = 1） | {stats['single']} | {pct_text(stats['single_ratio'])} |",
    "",
    "### 2) 重复次数分布（按 errorCount）",
    "",
    "| 分组 | 单元格数 | 占比 | 可视化 |",
    "|---|---:|---:|---|",
  ]
  for k in ["1次", "2-5次", "6-20次", "21-50次", "51次以上"]:
    v = stats["buckets"][k]
    p = (v / stats["dedup_rows"]) if stats["dedup_rows"] else 0.0
    lines.append(f"| {k} | {v} | {pct_text(p)} | `{bar(p)}` |")

  lines += [
    "",
    "### 3) Mismatch 原因汇总（按累计次数）",
    "",
    "| 原因 | 次数 | 占比 | 可视化 |",
    "|---|---:|---:|---|",
  ]
  for reason, cnt in stats["reasons"].most_common():
    p = (cnt / reason_total) if reason_total else 0.0
    lines.append(f"| {reason} | {cnt:,} | {pct_text(p)} | `{bar(p)}` |")

  lines += [
    "",
    "### 4) 高频重复 mismatch 单元格（Top 15）",
    "",
    "| 排名 | Sheet | Address | errorCount | 主要原因 | 首次出现Run | 最近出现Run |",
    "|---:|---|---|---:|---|---:|---:|",
  ]
  for i, row in enumerate(top_repeat, start=1):
    tr = row.get("topReasons", []) or []
    reason = tr[0].get("reason", "") if tr else ""
    sheet = sheet_alias(str(row.get("sheet", "")))
    lines.append(
      f"| {i} | {sheet} | {row.get('address', '')} | {int(row.get('errorCount', 0) or 0)} | {reason} | "
      f"{row.get('firstSeenRun', '')} | {row.get('lastSeenRun', '')} |"
    )

  lines += [
    "",
    "## 说明",
    "",
    "- 修改覆盖统计基于 `runs[].changedKeys`。",
    "- mismatch 原因统计基于 `batch_errors_dedup_100.json` 的 `topReasons` 聚合。",
    "- 本报告与HTML报告均为 UTF-8 编码。",
  ]
  return "\n".join(lines)


def bar_html(p: float) -> str:
  w = max(0.0, min(1.0, p)) * 100
  return (
    "<div class='bar-wrap'><div class='bar' style='width:"
    + f"{w:.2f}%"
    + ";'></div></div>"
  )


def render_html(stats: Dict[str, Any]) -> str:
  reason_total = sum(stats["reasons"].values())
  top_repeat = sorted(stats["rows"], key=lambda x: int(x.get("errorCount", 0) or 0), reverse=True)[:15]

  sheet_rows = []
  for sheet, cnt in stats["sheet_event"].most_common():
    e_pct = (cnt / stats["change_event_total"]) if stats["change_event_total"] else 0.0
    u_cnt = len(stats["sheet_unique"][sheet])
    u_pct = (u_cnt / stats["unique_change_cells"]) if stats["unique_change_cells"] else 0.0
    sheet_rows.append(
      "<tr>"
      + f"<td>{html.escape(sheet)}</td><td>{cnt:,}</td><td>{pct_text(e_pct)}</td>"
      + f"<td>{u_cnt}</td><td>{pct_text(u_pct)}</td><td>{bar_html(e_pct)}</td>"
      + "</tr>"
    )

  bucket_rows = []
  for k in ["1次", "2-5次", "6-20次", "21-50次", "51次以上"]:
    v = stats["buckets"][k]
    p = (v / stats["dedup_rows"]) if stats["dedup_rows"] else 0.0
    bucket_rows.append(
      "<tr>"
      + f"<td>{k}</td><td>{v}</td><td>{pct_text(p)}</td><td>{bar_html(p)}</td>"
      + "</tr>"
    )

  reason_rows = []
  for reason, cnt in stats["reasons"].most_common():
    p = (cnt / reason_total) if reason_total else 0.0
    reason_rows.append(
      "<tr>"
      + f"<td>{html.escape(reason)}</td><td>{cnt:,}</td><td>{pct_text(p)}</td><td>{bar_html(p)}</td>"
      + "</tr>"
    )

  top_rows = []
  for i, row in enumerate(top_repeat, start=1):
    tr = row.get("topReasons", []) or []
    reason = tr[0].get("reason", "") if tr else ""
    sheet = sheet_alias(str(row.get("sheet", "")))
    top_rows.append(
      "<tr>"
      + f"<td>{i}</td><td>{html.escape(sheet)}</td><td>{html.escape(str(row.get('address', '')))}</td>"
      + f"<td>{int(row.get('errorCount', 0) or 0)}</td><td>{html.escape(reason)}</td>"
      + f"<td>{row.get('firstSeenRun', '')}</td><td>{row.get('lastSeenRun', '')}</td>"
      + "</tr>"
    )

  return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Hyperformula 批量测试汇报（100批）</title>
  <style>
    :root {{
      --bg: #f6f8fb;
      --card: #ffffff;
      --text: #1c2430;
      --muted: #5e6b7a;
      --line: #d9e1ea;
      --primary: #1f77b4;
      --primary-soft: #dcecff;
      --ok: #2e8b57;
      --warn: #c0392b;
    }}
    body {{
      margin: 0;
      font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      background: linear-gradient(180deg, #eef3f9 0%, var(--bg) 100%);
      color: var(--text);
    }}
    .wrap {{
      max-width: 1200px;
      margin: 28px auto;
      padding: 0 16px 40px;
    }}
    h1, h2, h3 {{ margin: 0 0 12px; }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 12px;
      margin: 14px 0 18px;
    }}
    .kpi {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 14px 16px;
    }}
    .kpi .label {{ color: var(--muted); font-size: 13px; }}
    .kpi .value {{ font-size: 26px; font-weight: 700; margin-top: 4px; }}
    .section {{
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 16px;
      margin-top: 14px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
      margin-top: 8px;
    }}
    th, td {{
      border-bottom: 1px solid var(--line);
      text-align: left;
      padding: 8px 10px;
      vertical-align: middle;
    }}
    th {{
      color: var(--muted);
      font-weight: 600;
      background: #fafcff;
    }}
    .bar-wrap {{
      width: 160px;
      height: 10px;
      border-radius: 999px;
      background: var(--primary-soft);
      overflow: hidden;
    }}
    .bar {{
      height: 10px;
      background: var(--primary);
    }}
    .note {{ color: var(--muted); font-size: 13px; }}
    .ok {{ color: var(--ok); font-weight: 600; }}
    .warn {{ color: var(--warn); font-weight: 600; }}
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Hyperformula 批量测试汇报（100批）</h1>
    <p>一眼看懂：共执行 <b>{stats['run_count']}</b> 批，Compared <b>{stats['compared']:,}</b>，Matched <span class="ok">{stats['matched']:,}</span>，Mismatched <span class="warn">{stats['mismatched']:,}</span>，匹配率 <b>{pct_text(stats['match_rate'])}</b>。</p>
    <div class="grid">
      <div class="kpi"><div class="label">测试批次</div><div class="value">{stats['run_count']}</div></div>
      <div class="kpi"><div class="label">Compared</div><div class="value">{stats['compared']:,}</div></div>
      <div class="kpi"><div class="label">Matched</div><div class="value">{stats['matched']:,}</div></div>
      <div class="kpi"><div class="label">Mismatched</div><div class="value">{stats['mismatched']:,}</div></div>
      <div class="kpi"><div class="label">匹配率</div><div class="value">{pct_text(stats['match_rate'])}</div></div>
      <div class="kpi"><div class="label">去重 mismatch 单元格</div><div class="value">{stats['dedup_rows']}</div></div>
    </div>

    <div class="section">
      <h2>（1）修改单元格范围</h2>
      <p>修改事件总数 <b>{stats['change_event_total']:,}</b>，唯一改单元格 <b>{stats['unique_change_cells']:,}</b>，覆盖表格 <b>{stats['covered_sheets']}</b>。</p>
      <table>
        <thead><tr><th>表格</th><th>修改事件数</th><th>事件占比</th><th>唯一改单元格数</th><th>唯一占比</th><th>事件占比可视化</th></tr></thead>
        <tbody>
          {''.join(sheet_rows)}
        </tbody>
      </table>
    </div>

    <div class="section">
      <h2>（2）Mismatch 重复情况</h2>
      <p>去重后 mismatch 单元格 <b>{stats['dedup_rows']}</b>，其中重复出现（errorCount&gt;1）<b>{stats['repeated']}</b>，占比 <b>{pct_text(stats['repeat_ratio'])}</b>。</p>
      <table>
        <thead><tr><th>分组</th><th>单元格数</th><th>占比</th><th>可视化</th></tr></thead>
        <tbody>
          {''.join(bucket_rows)}
        </tbody>
      </table>
    </div>

    <div class="section">
      <h2>（3）Mismatch 原因汇总</h2>
      <table>
        <thead><tr><th>原因</th><th>次数</th><th>占比</th><th>可视化</th></tr></thead>
        <tbody>
          {''.join(reason_rows)}
        </tbody>
      </table>
    </div>

    <div class="section">
      <h2>（4）高频重复 mismatch 单元格（Top 15）</h2>
      <table>
        <thead><tr><th>排名</th><th>Sheet</th><th>Address</th><th>errorCount</th><th>主要原因</th><th>首次Run</th><th>最近Run</th></tr></thead>
        <tbody>
          {''.join(top_rows)}
        </tbody>
      </table>
    </div>

    <p class="note">本文件为 UTF-8 编码。若显示异常，请确认浏览器或编辑器按 UTF-8 打开。</p>
  </div>
</body>
</html>"""


def main() -> None:
  parser = argparse.ArgumentParser(description="Render batch report markdown/html from summary + dedup error JSON.")
  parser.add_argument("--summary", default="logs/batch_summary_100.json")
  parser.add_argument("--errors", default="logs/batch_errors_dedup_100.json")
  parser.add_argument("--md-out", default="logs/batch_report_100_final.md")
  parser.add_argument("--html-out", default="logs/batch_report_100_final.html")
  args = parser.parse_args()

  summary = json.loads(Path(args.summary).resolve().read_text(encoding="utf-8-sig"))
  dedup = json.loads(Path(args.errors).resolve().read_text(encoding="utf-8-sig"))
  stats = build_stats(summary, dedup)

  md_out = Path(args.md_out).resolve()
  html_out = Path(args.html_out).resolve()
  md_out.parent.mkdir(parents=True, exist_ok=True)
  html_out.parent.mkdir(parents=True, exist_ok=True)

  md_out.write_text(render_markdown(stats), encoding="utf-8")
  html_out.write_text(render_html(stats), encoding="utf-8")

  print(f"Markdown: {md_out}")
  print(f"HTML: {html_out}")


if __name__ == "__main__":
  main()
