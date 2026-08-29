from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any

from .models import AuditReport


def write_json_report(report: AuditReport, path: Path) -> None:
    payload = report.model_dump(mode="json", exclude={"differences"})
    keys = sorted([*payload, "differences"])
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write("{\n")
        for index, key in enumerate(keys):
            handle.write(f"  {json.dumps(key, ensure_ascii=False)}: ")
            if key == "differences":
                handle.write("[")
                for difference_index, difference in enumerate(report.differences):
                    if difference_index:
                        handle.write(",")
                    handle.write("\n    ")
                    handle.write(json.dumps(difference.model_dump(mode="json"), ensure_ascii=False, sort_keys=True, default=str))
                if report.differences:
                    handle.write("\n  ")
                handle.write("]")
            else:
                handle.write(json.dumps(payload[key], ensure_ascii=False, indent=2, sort_keys=True, default=str).replace("\n", "\n  "))
            handle.write(",\n" if index + 1 < len(keys) else "\n")
        handle.write("}\n")


def write_html_report(report: AuditReport, path: Path) -> None:
    summary = report.summary
    header = f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>Excel 核验报告</title>
<style>body{{font-family:system-ui,sans-serif;margin:32px;color:#222}}table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccc;padding:6px;text-align:left}}th{{background:#f5f5f5}}.summary{{display:flex;gap:24px;flex-wrap:wrap}}</style></head>
<body><h1>Excel 核验报告</h1><p>任务：{html.escape(report.job_id)} · 规则：{html.escape(report.schema_id)}@{html.escape(report.schema_version)}</p>
<p>标准快照：{html.escape(report.standard_snapshot_id)} · SHA-256：{html.escape(report.standard_sha256)}</p>
<div class="summary"><span>匹配记录：{summary.matched_records}</span><span>多余记录：{summary.extra_records}</span><span>缺失记录：{summary.missing_records}</span><span>差异：{summary.differences}</span></div>
<h2>工作簿结构</h2><pre>{html.escape(json.dumps(report.workbook_structure, ensure_ascii=False, indent=2))}</pre>
<h2>字段统计</h2><pre>{html.escape(json.dumps(report.field_statistics, ensure_ascii=False, indent=2))}</pre>
<h2>差异明细</h2><table><thead><tr><th>类型</th><th>级别</th><th>工作表</th><th>单元格</th><th>字段</th><th>业务主键</th><th>Excel 原值</th><th>Excel 规范值</th><th>标准原值</th><th>标准规范值</th><th>规则 ID</th><th>动作</th><th>修复状态</th><th>说明</th></tr></thead><tbody>"""
    with path.open("w", encoding="utf-8", newline="\n") as handle:
        handle.write(header)
        for item in report.differences:
            handle.write("<tr>" + "".join(f"<td>{html.escape(str(value if value is not None else ''))}</td>" for value in [
                item.type.value, item.severity, item.sheet_name, item.cell, item.canonical_field,
                json.dumps(item.business_key, ensure_ascii=False, sort_keys=True) if item.business_key else "",
                item.excel_raw_value, item.excel_normalized_value, item.standard_raw_value,
                item.standard_normalized_value, item.rule_id, item.render_action, item.repair_status, item.message,
            ]) + "</tr>")
        handle.write("</tbody></table></body></html>")
