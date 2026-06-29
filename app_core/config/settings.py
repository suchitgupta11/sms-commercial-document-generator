from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "SMS Commercial Document Generator"
APP_VERSION = "1.0.0"
COMPANY_NAME = "SMS Controls & Automation"
APP_TITLE = f"{APP_NAME} v{APP_VERSION}"

PRIMARY = "#0B2F4F"
SECONDARY = "#155C94"
LIGHT = "#EAF4FB"
BG = "#F5F8FB"
TEXT = "#172B4D"
WHITE = "#FFFFFF"
SUCCESS = "#0F766E"
DANGER = "#B42318"
BORDER = "#D6DEE6"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_base_dir() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)).resolve()
    return Path(__file__).resolve().parents[2]


def executable_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).parent.resolve()
    return Path(__file__).resolve().parents[2]


def user_data_dir() -> Path:
    if sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support" / APP_NAME
    elif os.name == "nt":
        root = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA") or str(Path.home() / "AppData" / "Local")
        base = Path(root) / APP_NAME
    else:
        base = Path(os.getenv("XDG_DATA_HOME", Path.home() / ".local" / "share")) / APP_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base


def user_output_dir() -> Path:
    path = Path.home() / "Documents" / "SMS Commercial Documents"
    path.mkdir(parents=True, exist_ok=True)
    return path


def logs_dir() -> Path:
    path = user_data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def temp_dir() -> Path:
    path = user_data_dir() / "temp"
    path.mkdir(parents=True, exist_ok=True)
    return path

BASE_DIR = app_base_dir()
APP_DIR = BASE_DIR / "app"
UI_DIR = BASE_DIR / "ui"
APP_CORE_DIR = BASE_DIR / "app_core"
MODULES_DIR = APP_CORE_DIR / "engines" / "modules"
RESOURCES_DIR = BASE_DIR / "resources"
TEMPLATES_DIR = RESOURCES_DIR / "templates"
ASSETS_DIR = RESOURCES_DIR / "assets"
LEGACY_ASSETS_DIR = BASE_DIR / "assets"
OUTPUT_DIR = user_output_dir()
TEMP_DIR = temp_dir()
LOG_DIR = logs_dir()
