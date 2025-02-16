from datetime import datetime
import logging.config
from pathlib import Path
from typing import Any


def setup_logging(log_folder: str | None = None) -> None:
    log_folder = log_folder or "./logs"
    Path(log_folder).mkdir(exist_ok=True)

    config: dict[str, Any] = {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {
            "default": {
                "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            },
        },
        "handlers": {
            "console": {
                "class": "logging.StreamHandler",
                "formatter": "default",
            },
            "file": {
                "class": "logging.FileHandler",
                "filename": f"{log_folder}/{datetime.now().strftime('%Y%m%d%H%M%S')}.log",
                "formatter": "default",
            },
        },
        "loggers": {
            "": {
                "handlers": ["console", "file"],
                "level": "DEBUG",
            },
        },
    }

    logging.config.dictConfig(config)
