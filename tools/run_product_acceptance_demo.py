from __future__ import annotations

import argparse
import json
import shutil
from decimal import Decimal
from pathlib import Path

from excel_auditor.models import FieldType, ValidationConfig
from excel_auditor.product_workflow import (
    CatalogFieldDefinition,
    CatalogFieldSource,
    CategoryDefinition,
    InMemoryCatalogAdapter,
)
from excel_auditor.product_workflow.service import ProductWorkflowService
from excel_auditor.rules import load_rules
from excel_auditor.service import AuditService


def platform_field(
    field_id: str,
    title: str,
    *,
    source: CatalogFieldSource = CatalogFieldSource.PLATFORM_ATTRIBUTE,
    field_type: FieldType = FieldType.STRING,
    required: bool = False,
    multiple: bool = False,
    order: int = 0,
    enum_values: list[str] | None = None,
    number_format: str | None = None,
    timezone: str | None = None,
    validation: ValidationConfig | None = None,
) -> CatalogFieldDefinition:
    return CatalogFieldDefinition(
        field_id=field_id,
        title=title,
        source=source,
        field_type=field_type,
        required=required,
        multiple=multiple,
        display_order=order,
        category_id="cat-phone",
        attribute_id=field_id,
        enum_values=enum_values or [],
        number_format=number_format,
        timezone=timezone,
        validation=validation or ValidationConfig(),
    )


def demo_catalog() -> InMemoryCatalogAdapter:
    phone_fields = [
        platform_field("brand", "品牌", required=True, order=10),
        platform_field("model", "型号", required=True, order=20),
        platform_field(
            "price",
            "销售价",
            field_type=FieldType.DECIMAL,
            required=True,
            order=30,
            number_format="#,##0.00",
            validation=ValidationConfig(min=Decimal("0"), max=Decimal("1000000")),
        ),
        platform_field(
            "listing_at",
            "上架时间",
            field_type=FieldType.DATETIME,
            order=40,
            number_format="yyyy-mm-dd hh:mm",
            timezone="Asia/Shanghai",
        ),
        platform_field(
            "color",
            "颜色",
            source=CatalogFieldSource.PLATFORM_SPECIFICATION,
            field_type=FieldType.ENUM,
            required=True,
            multiple=True,
            order=10,
            enum_values=["黑色", "白色", "蓝色"],
        ),
        platform_field(
            "storage",
            "存储容量",
            source=CatalogFieldSource.PLATFORM_SPECIFICATION,
            field_type=FieldType.ENUM,
            required=True,
            multiple=True,
            order=20,
            enum_values=["128GB", "256GB", "512GB"],
        ),
    ]
    return InMemoryCatalogAdapter(
        [
            CategoryDefinition(
                category_id="cat-phone",
                name="手机",
                path=["数码", "手机"],
                aliases=["智能手机", "手机数码"],
            ),
            CategoryDefinition(
                category_id="cat-shoe",
                name="运动鞋",
                path=["服饰", "鞋靴", "运动鞋"],
                aliases=["跑鞋"],
            ),
        ],
        {"cat-phone": phone_fields, "cat-shoe": []},
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Run a deterministic product-normalization acceptance demo")
    parser.add_argument("input", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()

    output_dir = args.output_dir.resolve()
    runtime = output_dir / "runtime"
    rules = load_rules(Path("configs/examples/product-normalization.yaml"))
    audit = AuditService(runtime)
    workflow = ProductWorkflowService(audit)
    job_id = audit.create_job(user_id="acceptance-demo")
    job_input = audit.job_directory(job_id) / "product-input.xlsx"
    shutil.copy2(args.input.resolve(), job_input)

    workflow.run(job_id, job_input, rules, demo_catalog(), actor_id="acceptance-demo")
    status = audit.status(job_id)
    if status["status"] not in {"completed", "manual_review"}:
        raise RuntimeError(json.dumps(status, ensure_ascii=False, indent=2))

    job_dir = audit.job_directory(job_id)
    artifact_names = status.get("artifacts", {})
    exported: dict[str, str] = {}
    for key in ("product_excel", "product_result", "product_issues", "product_manifest"):
        name = artifact_names.get(key)
        if not name:
            continue
        source = job_dir / name
        suffix = source.suffix
        destination = output_dir / f"模拟商品-{key}{suffix}"
        shutil.copy2(source, destination)
        exported[key] = str(destination)

    summary = {
        "job_id": job_id,
        "status": status["status"],
        "category_count": status.get("category_count"),
        "unresolved_row_count": status.get("unresolved_row_count"),
        "review_count": status.get("review_count"),
        "issue_count": status.get("issue_count"),
        "output_sha256": status.get("output_sha256"),
        "artifacts": exported,
    }
    (output_dir / "模拟商品-验收摘要.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
