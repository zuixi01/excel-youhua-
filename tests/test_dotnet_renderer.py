import hashlib
import json
import os
import subprocess
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table
from openpyxl.workbook.defined_name import DefinedName

from excel_auditor.models import RuleSet
from excel_auditor.rendering import DotNetOpenXmlRenderer
from excel_auditor.service import AuditService


def test_dotnet_renderer_self_check_contract():
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    assert DotNetOpenXmlRenderer(Path(command)).self_check() == "0.1.0"


def test_dotnet_renderer_marks_sparse_row_across_used_columns(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source, output, manifest_path = tmp_path / "sparse.xlsx", tmp_path / "sparse-output.xlsx", tmp_path / "manifest.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Name", "Amount"])
    sheet.append(["E1"])
    book.save(source)
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0",
        "job_id": "job_sparse_row",
        "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "operations": [{"type": "mark_row", "sheet": "Data", "row": 2, "fill_color": "F4CCCC", "comment": "extra"}],
    }), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    rendered = load_workbook(output)
    assert all(rendered["Data"].cell(2, column).fill.fgColor.rgb.endswith("F4CCCC") for column in range(1, 4))
    assert rendered["Data"].calculate_dimension() == "A1:C2"
    rendered.close()


def test_dotnet_renderer_validates_hash_marks_and_inserts(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source = tmp_path / "input.xlsx"
    output = tmp_path / "output.xlsx"
    manifest_path = tmp_path / "manifest.json"
    report_path = tmp_path / "report.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员信息"
    sheet.append(["员工编号", "工资"])
    sheet.append(["E001", "99"])
    book.save(source)
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    report_path.write_text(json.dumps({"summary": {"matched_records": 1}, "differences": [{"type": "VALUE_MISMATCH", "sheet_name": "人员信息", "cell": "B2", "canonical_field": "salary", "message": "值不一致"}]}, ensure_ascii=False), encoding="utf-8")
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0",
        "job_id": "job_contract",
        "input_sha256": digest,
        "operations": [
            {"type": "mark_cell", "sheet": "人员信息", "cell": "B2", "fill_color": "FFE599", "comment": "值不一致"},
            {"type": "insert_column", "sheet": "人员信息", "before": "B", "canonical_field": "name", "header_row": 1, "header_value": "姓名", "fill_color": "D9EAD3"},
            {"type": "add_or_replace_report_sheet", "name": "核验报告", "source_json": "report.json"},
        ],
    }, ensure_ascii=False), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    result = json.loads(completed.stdout)
    assert result["success"] is True
    assert {item["difference_id"] for item in result["operation_results"] if item["difference_id"]} == set()
    rendered = load_workbook(output)
    assert rendered["人员信息"]["B1"].value == "姓名"
    assert rendered["人员信息"]["C2"].value == "99"
    assert rendered["人员信息"]["C2"].fill.fgColor.rgb.endswith("FFE599")
    assert rendered["人员信息"]["C2"].comment is not None
    assert "值不一致" in rendered["人员信息"]["C2"].comment.text
    assert rendered["核验报告"]["A2"].value == "matched_records"
    first_style_count = len(rendered._cell_styles)
    first_fill_count = len(rendered._fills)
    rendered.close()

    repeated = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert repeated.returncode == 0, repeated.stderr
    assert json.loads(repeated.stdout)["success"] is True
    rerendered = load_workbook(output)
    assert [sheet_name for sheet_name in rerendered.sheetnames if sheet_name == "核验报告"] == ["核验报告"]
    assert [sheet_name for sheet_name in rerendered.sheetnames if sheet_name == "__ExcelAuditorMetadata"] == ["__ExcelAuditorMetadata"]
    assert rerendered["人员信息"]["B1"].value == "姓名"
    assert rerendered["人员信息"]["C2"].comment.text == "值不一致"
    assert len(rerendered._cell_styles) == first_style_count
    assert len(rerendered._fills) == first_fill_count
    rerendered.close()

    dry_output = tmp_path / "dry-run.xlsx"
    dry_output.write_bytes(b"pre-existing-output")
    dry = subprocess.run([command, "--input", str(source), "--output", str(dry_output), "--manifest", str(manifest_path), "--dry-run"], capture_output=True, text=True, check=False)
    assert dry.returncode == 0, dry.stderr
    dry_result = json.loads(dry.stdout)
    assert dry_result["success"] is True and dry_result["dry_run"] is True
    assert dry_output.read_bytes() == b"pre-existing-output"


def test_dotnet_renderer_writes_typed_cells_formats_validation_metadata_and_rejects_unknown_ops(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source, output, manifest_path = tmp_path / "typed.xlsx", tmp_path / "typed-output.xlsx", tmp_path / "manifest.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Amount"])
    sheet.append(["E1", 1])
    sheet.add_table(Table(displayName="TypedTable", ref="A1:B2"))
    book.save(source)
    manifest = {
        "manifest_version": "1.0",
        "job_id": "job_typed",
        "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "metadata": {
            "schema_id": "typed",
            "schema_version": "1.0.0",
            "schema_sha256": "a" * 64,
            "standard_snapshot_id": "std_example",
            "standard_sha256": "b" * 64,
            "result_sha256": "c" * 64,
        },
        "operations": [
            {"type": "set_cell", "sheet": "Data", "cell": "B2", "value": "12.50", "field_type": "decimal", "number_format": "0.00", "fill_color": "D9EAD3"},
            {"type": "insert_column", "sheet": "Data", "before": "B", "canonical_field": "score", "header_row": 1, "header_value": "Score", "fill_color": "D9EAD3", "field_type": "decimal", "number_format": "0.00", "formula_template": "=ROW()", "validation": {"type": "decimal", "min": "0", "max": "100", "allow_blank": False}},
            {"type": "append_row", "sheet": "Data", "row": 3, "values": [{"cell": "A3", "value": "E2", "field_type": "string"}, {"cell": "C3", "value": "9.75", "field_type": "decimal", "number_format": "0.00"}], "fill_color": "D9EAD3"},
        ],
    }
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    rendered = load_workbook(output, data_only=False)
    data = rendered["Data"]
    assert data["B2"].value == "=ROW()"
    assert data["B2"].number_format == "0.00"
    assert data["C2"].value == 12.5 and data["C2"].data_type == "n"
    assert data["C3"].value == 9.75 and data["C3"].data_type == "n"
    assert data["C3"].number_format == "0.00"
    assert data.tables["TypedTable"].ref == "A1:C3"
    assert str(data.data_validations.dataValidation[0].sqref) == "B2:B3"
    assert data.calculate_dimension() == "A1:C3"
    metadata = rendered["__ExcelAuditorMetadata"]
    assert metadata.sheet_state == "veryHidden"
    assert metadata["B2"].value == "job_typed"
    assert metadata["A9"].value == "result_content_sha256"
    assert metadata["B9"].value == json.loads(completed.stdout)["result_content_sha256"]
    assert len(metadata["B9"].value) == 64
    rendered.close()

    manifest["operations"] = [{"type": "unrecognized_operation", "sheet": "Data"}]
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    rejected = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert rejected.returncode != 0
    assert json.loads(rejected.stderr)["error_code"] == "MANIFEST_OR_STRUCTURE_INVALID"


def test_dotnet_renderer_updates_table_validation_defined_name_and_formula_ranges(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source, output, manifest_path = tmp_path / "structure.xlsx", tmp_path / "structure-output.xlsx", tmp_path / "manifest.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "结构"
    sheet.append(["编号", "数量", "单价", "合计", "文本引用"])
    sheet.append(["E001", 2, 5, "=B2*C2", '=IF(B2="B1",C2,0)'])
    sheet.append(["E002", 3, 4, "=B3*C3", '=IF(B3="C2",C3,0)'])
    sheet.add_table(Table(displayName="DataTable", ref="A1:C3"))
    validation = DataValidation(type="whole", operator="greaterThan", formula1="0")
    validation.add("C2:C3")
    sheet.add_data_validation(validation)
    sheet.auto_filter.ref = "A1:E3"
    sheet.auto_filter.add_filter_column(2, ["5"])
    sheet.freeze_panes = "C2"
    sheet.sheet_view.selection[0].activeCell = "C2"
    sheet.sheet_view.selection[0].sqref = "C2:D3"
    sheet["A3"]._hyperlink = Hyperlink(ref="A3", location="'结构'!B2", display="jump")
    book.defined_names.add(DefinedName("DataRange", attr_text="'结构'!$A$1:$C$3"))
    book.defined_names.add(DefinedName("LocalRange", attr_text="$A$1:$C$3", localSheetId=0))
    other = book.create_sheet("汇总")
    other["A1"] = "=结构!B2+'结构'!C2"
    other["A2"] = "=B1"
    sheet["G2"] = "=汇总!B1+结构!B2"
    book.save(source)
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0", "job_id": "job_ranges", "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "operations": [{"type": "insert_column", "sheet": "结构", "before": "B", "canonical_field": "department", "header_row": 1, "header_value": "部门", "fill_color": "D9EAD3"}],
    }, ensure_ascii=False), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    rendered = load_workbook(output)
    result_sheet = rendered["结构"]
    assert result_sheet.tables["DataTable"].ref == "A1:D3"
    assert [column.name for column in result_sheet.tables["DataTable"].tableColumns] == ["编号", "部门", "数量", "单价"]
    assert str(result_sheet.data_validations.dataValidation[0].sqref) == "D2:D3"
    assert result_sheet.auto_filter.ref == "A1:F3"
    assert result_sheet.auto_filter.filterColumn[0].colId == 3
    assert result_sheet.freeze_panes == "D2"
    assert result_sheet.sheet_view.selection[0].activeCell == "D2"
    assert str(result_sheet.sheet_view.selection[0].sqref) == "D2:E3"
    assert result_sheet["A3"].hyperlink.location == "'结构'!C2"
    assert rendered.defined_names["DataRange"].attr_text == "'结构'!$A$1:$D$3"
    assert result_sheet.defined_names["LocalRange"].attr_text == "$A$1:$D$3"
    assert result_sheet["E2"].value == "=C2*D2"
    assert result_sheet["F2"].value == '=IF(C2="B1",D2,0)'
    assert result_sheet["F3"].value == '=IF(C3="C2",D3,0)'
    assert result_sheet["H2"].value == "=汇总!B1+结构!C2"
    assert rendered["汇总"]["A1"].value == "=结构!C2+'结构'!D2"
    assert rendered["汇总"]["A2"].value == "=B1"
    rendered.close()


def test_dotnet_renderer_supports_insert_after_contract(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source, output, manifest_path = tmp_path / "after.xlsx", tmp_path / "after-output.xlsx", tmp_path / "manifest.json"
    book = Workbook(); sheet = book.active; sheet.title = "Data"; sheet.append(["A", "B"]); sheet.append([1, 2]); book.save(source)
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0", "job_id": "job_after", "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "operations": [{"type": "insert_column", "sheet": "Data", "after": "B", "canonical_field": "c", "header_row": 1, "header_value": "C", "fill_color": "D9EAD3"}],
    }), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    rendered = load_workbook(output)
    assert [rendered["Data"].cell(1, index).value for index in range(1, 4)] == ["A", "B", "C"]
    rendered.close()


def test_dotnet_renderer_rejects_unsafe_complex_formula_insertions(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    source, output, manifest_path = tmp_path / "complex.xlsx", tmp_path / "complex-output.xlsx", tmp_path / "manifest.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Amount", "Total"])
    sheet.append(["E1", 1, "=SUM(B:B)"])
    book.save(source)
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0", "job_id": "job_complex", "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "operations": [{"type": "insert_column", "sheet": "Data", "before": "B", "canonical_field": "score", "header_row": 1, "header_value": "Score", "fill_color": "D9EAD3"}],
    }), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode != 0
    failure = json.loads(completed.stderr)
    assert failure["error_code"] == "UNSUPPORTED_FEATURE"
    assert "complex" in failure["message"]
    assert not output.exists()


def test_dotnet_renderer_preserves_macro_and_signature_parts_during_real_edit(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    if not command:
        pytest.skip("set EXCEL_RENDERER_COMMAND to run the .NET renderer contract test")
    xlsx, source, output, manifest_path = tmp_path / "base.xlsx", tmp_path / "macro.xlsm", tmp_path / "macro-output.xlsm", tmp_path / "manifest.json"
    book = Workbook()
    book.active.title = "Data"
    book.active.append(["ID"])
    book.active.append(["E1"])
    book.save(xlsx)
    macro_payload = b"\x00VBA-PROJECT-GOLDEN\xff\x10"
    signature_payload = b"\x00VBA-SIGNATURE-GOLDEN\xfe\x20"
    with zipfile.ZipFile(xlsx) as incoming, zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for info in incoming.infolist():
            data = incoming.read(info.filename)
            if info.filename == "[Content_Types].xml":
                text = data.decode("utf-8").replace(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                ).replace(
                    "</Types>",
                    '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                    '<Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/>'
                    "</Types>",
                )
                data = text.encode("utf-8")
            elif info.filename == "xl/_rels/workbook.xml.rels":
                text = data.decode("utf-8").replace("</Relationships>", '<Relationship Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin" Id="rIdVba"/></Relationships>')
                data = text.encode("utf-8")
            outgoing.writestr(info, data)
        outgoing.writestr("xl/vbaProject.bin", macro_payload)
        outgoing.writestr("xl/vbaProjectSignature.bin", signature_payload)
        outgoing.writestr(
            "xl/_rels/vbaProject.bin.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdSignature" '
            'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" '
            'Target="vbaProjectSignature.bin"/>'
            '</Relationships>',
        )
    manifest_path.write_text(json.dumps({
        "manifest_version": "1.0", "job_id": "job_macro", "input_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "operations": [{
            "type": "mark_cell", "sheet": "Data", "cell": "A2", "fill_color": "FFF2CC",
            "comment": "audited", "difference_id": "diff_macro",
        }],
    }), encoding="utf-8")
    completed = subprocess.run([command, "--input", str(source), "--output", str(output), "--manifest", str(manifest_path)], capture_output=True, text=True, check=False)
    assert completed.returncode == 0, completed.stderr
    with zipfile.ZipFile(output) as archive:
        assert archive.read("xl/vbaProject.bin") == macro_payload
        assert archive.read("xl/vbaProjectSignature.bin") == signature_payload
    rendered = load_workbook(output, keep_vba=True)
    assert rendered["Data"]["A2"].comment.text == "audited"
    assert rendered["Data"]["A2"].fill.fgColor.rgb.endswith("FFF2CC")
    rendered.close()


def test_rendered_output_opens_and_resaves_in_libreoffice(tmp_path):
    command = os.environ.get("EXCEL_RENDERER_COMMAND")
    libreoffice = os.environ.get("LIBREOFFICE_COMMAND")
    if not command or not libreoffice:
        pytest.skip("set renderer and LibreOffice commands to run compatibility test")
    rules = RuleSet.model_validate({
        "schema_id": "compat", "schema_version": "1.0.0", "name": "Compatibility",
        "sheets": [{"id": "data", "name": "Data", "primary_key": ["id"], "columns": [
            {"name": "id", "title": "ID", "required": True},
            {"name": "name", "title": "Name", "required": True},
        ]}],
    })
    source, standard = tmp_path / "compat.xlsx", tmp_path / "standard.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["ID", "Name"])
    sheet.append(["E1", "Alice"])
    book.save(source)
    standard.write_text(json.dumps({"data": [{"id": "E1", "name": "Alice"}]}), encoding="utf-8")
    service = AuditService(tmp_path / "runtime")
    job_id = service.create_job()
    service.run(job_id, source, standard, rules)
    output = service.artifact(job_id, "excel")
    converted = tmp_path / "libreoffice"
    converted.mkdir()
    completed = subprocess.run([libreoffice, "--headless", "--convert-to", "xlsx", "--outdir", str(converted), str(output)], capture_output=True, text=True, timeout=120, check=False)
    assert completed.returncode == 0, completed.stderr
    resaved = converted / output.with_suffix(".xlsx").name
    assert resaved.is_file()
    load_workbook(resaved, read_only=True).close()
