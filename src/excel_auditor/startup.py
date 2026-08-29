from __future__ import annotations

import os
from pathlib import Path

import uvicorn
from alembic import command
from alembic.config import Config

from .security_config import validate_production_environment


def run() -> None:
    validate_production_environment()
    config_path = Path(os.environ.get("EXCEL_AUDITOR_ALEMBIC_CONFIG", Path.cwd() / "alembic.ini"))
    if os.environ.get("DATABASE_URL"):
        config = Config(str(config_path))
        command.upgrade(config, "head")
    uvicorn.run("excel_auditor.api:app", host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
