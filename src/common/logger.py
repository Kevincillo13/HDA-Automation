import logging
import os

from src.common.config import get_settings
from src.common.run_context import get_run_id


def get_logger(name: str) -> logging.Logger:
    settings = get_settings()
    os.makedirs(settings.log_dir, exist_ok=True)

    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    log_path = os.path.join(
        settings.log_dir,
        f"{get_run_id()}.log",
    )

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger
