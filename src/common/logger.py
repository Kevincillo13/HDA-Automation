import logging
import os

from src.common.config import get_settings
from src.common.run_context import get_run_id


# Lista global de callbacks para la GUI
_gui_callbacks = []

class GuiLogHandler(logging.Handler):
    """Handler personalizado para enviar logs a la interfaz gráfica."""
    def emit(self, record):
        log_entry = self.format(record)
        for callback in _gui_callbacks:
            try:
                callback(log_entry)
            except:
                pass

def add_gui_callback(callback):
    """Registra una función para recibir actualizaciones de log."""
    if callback not in _gui_callbacks:
        _gui_callbacks.append(callback)

def get_logger(name: str) -> logging.Logger:
    settings = get_settings()
    os.makedirs(settings.log_dir, exist_ok=True)

    logger = logging.getLogger(name)
    if logger.handlers:
        # Verificar si ya tiene el handler de la GUI
        if not any(isinstance(h, GuiLogHandler) for h in logger.handlers):
            gui_handler = GuiLogHandler()
            formatter = logging.Formatter("%(asctime)s | %(message)s", datefmt="%H:%M:%S")
            gui_handler.setFormatter(formatter)
            logger.addHandler(gui_handler)
        return logger

    logger.setLevel(logging.INFO)
    run_id = get_run_id()
    log_path = os.path.join(settings.log_dir, f"{run_id}.log")

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )
    
    # Handler para archivo
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Handler para consola
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    # Handler para la GUI (terminal fresona)
    gui_handler = GuiLogHandler()
    gui_formatter = logging.Formatter("%(asctime)s | %(message)s", datefmt="%H:%M:%S")
    gui_handler.setFormatter(gui_formatter)
    logger.addHandler(gui_handler)

    return logger
