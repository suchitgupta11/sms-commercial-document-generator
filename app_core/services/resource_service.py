from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app_core.config.settings import ASSETS_DIR, RESOURCES_DIR, TEMPLATES_DIR, BASE_DIR


class ResourceManager:
    """Centralized resource resolver for normal Python and PyInstaller runtime."""

    @staticmethod
    def template(*parts: str) -> Path:
        return TEMPLATES_DIR.joinpath(*parts)

    @staticmethod
    def asset(*parts: str) -> Path:
        candidate = ASSETS_DIR.joinpath(*parts)
        if candidate.exists():
            return candidate
        return BASE_DIR.joinpath("assets", *parts)

    @staticmethod
    def branding(filename: str = "company.json") -> Path:
        return RESOURCES_DIR / "branding" / filename

    @staticmethod
    def read_json(path: Path, default: dict[str, Any] | None = None) -> dict[str, Any]:
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default or {}
