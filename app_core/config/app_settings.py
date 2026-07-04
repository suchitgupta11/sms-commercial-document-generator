from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from app_core.config.settings import user_data_dir
from app_core.config.theme import ENTERPRISE_DARK


@dataclass
class ApplicationPreferences:
    theme: str = ENTERPRISE_DARK.name
    auto_open_output_folder: bool = False
    auto_open_pdf_after_generate: bool = False
    remember_last_document_type: bool = True
    last_document_type: str = "Quotation"


class SettingsStore:
    def __init__(self, filename: str = "application_settings.json") -> None:
        self.path = user_data_dir() / "config" / filename
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> ApplicationPreferences:
        if not self.path.exists():
            prefs = ApplicationPreferences()
            self.save(prefs)
            return prefs
        try:
            raw: dict[str, Any] = json.loads(self.path.read_text(encoding="utf-8"))
            defaults = asdict(ApplicationPreferences())
            defaults.update(raw)
            return ApplicationPreferences(**defaults)
        except Exception:
            prefs = ApplicationPreferences()
            self.save(prefs)
            return prefs

    def save(self, prefs: ApplicationPreferences) -> None:
        self.path.write_text(json.dumps(asdict(prefs), indent=2), encoding="utf-8")
