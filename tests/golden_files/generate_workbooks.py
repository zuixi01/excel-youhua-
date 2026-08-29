"""Explicit maintainer tool for regenerating the committed Golden XLSX inputs."""

import json
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parent
PAYLOAD = json.loads((ROOT / "core_scenarios.json").read_text(encoding="utf-8"))
OUTPUT = ROOT / "workbooks"
OUTPUT.mkdir(exist_ok=True)

for scenario in PAYLOAD["scenarios"]:
    book = Workbook()
    sheet = book.active
    sheet.title = "员工"
    sheet.append(scenario["headers"])
    for row in scenario["rows"]:
        sheet.append(row)
    book.properties.created = datetime(2026, 8, 28, tzinfo=timezone.utc)
    book.properties.modified = datetime(2026, 8, 28, tzinfo=timezone.utc)
    destination = OUTPUT / f"{scenario['name']}.xlsx"
    temporary = destination.with_suffix(".tmp.xlsx")
    book.save(temporary)
    with zipfile.ZipFile(temporary) as source, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as target:
        for source_info in source.infolist():
            info = zipfile.ZipInfo(source_info.filename, date_time=(2026, 8, 28, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = source_info.external_attr
            data = source.read(source_info.filename)
            if source_info.filename == "docProps/core.xml":
                for field in (b"created", b"modified"):
                    pattern = rb"(<dcterms:" + field + rb"[^>]*>)[^<]*(</dcterms:" + field + rb">)"
                    data = re.sub(pattern, rb"\g<1>2026-08-28T00:00:00Z\g<2>", data)
            target.writestr(info, data, compresslevel=9)
    temporary.unlink()
