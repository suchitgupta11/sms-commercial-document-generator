from pathlib import Path

_ENGINE_MODULES = Path(__file__).resolve().parents[1] / "app_core" / "engines" / "modules"
if _ENGINE_MODULES.exists():
    __path__.append(str(_ENGINE_MODULES))
