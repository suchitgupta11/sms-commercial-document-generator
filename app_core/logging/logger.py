from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from app_core.config.settings import LOG_DIR, APP_NAME, APP_VERSION


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("sms_commercial_generator")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        log_file = Path(LOG_DIR) / "application.log"
        handler = RotatingFileHandler(log_file, maxBytes=1_000_000, backupCount=5, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(name)s | %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    logger.info("Starting %s v%s", APP_NAME, APP_VERSION)
    return logger
