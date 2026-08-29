import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from alembic import command
from alembic.config import Config
from sqlalchemy import create_engine, func, inspect, select
from sqlalchemy.orm import Session

from excel_auditor.persistence import AuditEventRow, ComparisonJobRow, DatabaseRepository, DifferenceIndexRow, SchemaVersionRow, StandardSnapshotRow
from excel_auditor.rules import load_rules
from excel_auditor.service import AuditService


def test_database_rule_registry_is_immutable_and_audited(tmp_path):
    database = DatabaseRepository(f"sqlite:///{tmp_path / 'audit.db'}")
    rules = load_rules(__import__("pathlib").Path("configs/examples/employee-roster.yaml"))
    database.publish_rule(rules, actor_id="tester")
    assert database.get_rule(rules.schema_id, rules.schema_version).content_sha256 == rules.content_sha256
    changed = rules.model_copy(update={"name": "changed"})
    try:
        database.publish_rule(changed)
        assert False, "expected immutable version conflict"
    except FileExistsError:
        pass
    with Session(database.engine) as session:
        assert session.scalar(select(func.count()).select_from(SchemaVersionRow)) == 1
        assert session.scalar(select(func.count()).select_from(AuditEventRow)) == 1


def test_database_rule_registry_is_tenant_isolated(tmp_path):
    database = DatabaseRepository(f"sqlite:///{tmp_path / 'tenant-rules.db'}")
    rules = load_rules(Path("configs/examples/employee-roster.yaml"))
    database.publish_rule(rules, tenant_id="tenant-a")
    database.publish_rule(rules, tenant_id="tenant-b")
    assert database.get_rule(rules.schema_id, rules.schema_version, "tenant-a").content_sha256 == rules.content_sha256
    try:
        database.get_rule(rules.schema_id, rules.schema_version, "tenant-c")
        assert False, "cross-tenant lookup must not resolve"
    except FileNotFoundError:
        pass


def test_service_persists_job_snapshot_and_difference_indexes(tmp_path):
    database = DatabaseRepository(f"sqlite:///{tmp_path / 'audit.db'}")
    rules = load_rules(Path("configs/examples/employee-roster.yaml"))
    excel = tmp_path / "input.xlsx"
    standard = tmp_path / "standard.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员信息"
    sheet.append(["标题"])
    sheet.append(["工号", "姓名", "工资", "部门"])
    sheet.append(["E001", "张三", "100", "技术部"])
    book.save(excel)
    standard.write_text(json.dumps({"employees": [{"employee_id": "E001", "employee_name": "张三", "salary": "200", "department": "技术部"}]}, ensure_ascii=False), encoding="utf-8")
    service = AuditService(tmp_path / "runtime", database=database)
    job_id = service.create_job()
    service.run(job_id, excel, standard, rules)
    second_job_id = service.create_job()
    service.run(second_job_id, excel, standard, rules)
    with Session(database.engine) as session:
        job = session.get(ComparisonJobRow, job_id)
        assert job is not None and job.status == "completed" and job.output_sha256
        assert session.scalar(select(func.count()).select_from(StandardSnapshotRow)) == 2
        assert session.scalar(select(func.count()).select_from(DifferenceIndexRow)) >= 2


def test_alembic_initial_migration_builds_expected_schema(tmp_path, monkeypatch):
    monkeypatch.delenv("DATABASE_URL", raising=False)
    database_path = tmp_path / "migrated.db"
    config = Config("alembic.ini")
    config.set_main_option("sqlalchemy.url", f"sqlite:///{database_path.as_posix()}")
    command.upgrade(config, "head")
    tables = set(inspect(create_engine(f"sqlite:///{database_path}")).get_table_names())
    assert {"schemas", "schema_versions", "standard_connections", "standard_snapshots", "comparison_jobs", "comparison_differences", "audit_events", "alembic_version"} <= tables


def test_expired_job_purge_removes_job_differences_and_unreferenced_snapshot_but_keeps_audit(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_AUDITOR_DELETE_DELAY_DAYS", "0")
    database = DatabaseRepository(f"sqlite:///{tmp_path / 'purge.db'}")
    rules = load_rules(Path("configs/examples/employee-roster.yaml"))
    excel, standard = tmp_path / "input.xlsx", tmp_path / "standard.json"
    book = Workbook()
    sheet = book.active
    sheet.title = "人员信息"
    sheet.append(["标题"])
    sheet.append(["工号", "姓名", "工资", "部门"])
    sheet.append(["E001", "张三", "100", "技术部"])
    book.save(excel)
    standard.write_text(json.dumps({"employees": [{"employee_id": "E001", "employee_name": "张三", "salary": "200", "department": "技术部"}]}, ensure_ascii=False), encoding="utf-8")
    service = AuditService(tmp_path / "runtime", database=database)
    job_id = service.create_job(tenant_id="tenant-a", user_id="tester")
    service.run(job_id, excel, standard, rules)
    service.soft_delete(job_id, "tester")
    assert service.purge_expired(datetime.now(timezone.utc)) == [job_id]
    with Session(database.engine) as session:
        assert session.get(ComparisonJobRow, job_id) is None
        assert session.scalar(select(func.count()).select_from(DifferenceIndexRow)) == 0
        assert session.scalar(select(func.count()).select_from(StandardSnapshotRow)) == 0
        assert session.scalar(select(func.count()).select_from(AuditEventRow)) > 0
