import os
import queue
import sys
import threading
import time

import webview

from src.common.config import get_settings
from src.common.logger import add_gui_callback, get_logger
from src.common.settings_manager import SettingsManager
from src.common.system import kill_processes
from src.hda_web.ticket_processing import process_all_tickets


def get_resource_path(relative_path):
    """Obtiene la ruta absoluta para recursos, compatible con PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


class ProcessorAPI:
    """API expuesta a JavaScript para controlar la automatización."""

    def __init__(self):
        self.cancel_event = threading.Event()
        self.is_running = False
        self.window = None
        self.settings_manager = SettingsManager()
        self.worker_thread: threading.Thread | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.state_lock = threading.Lock()
        self.final_status: dict[str, str | bool] | None = None
        self.shutdown_started = threading.Event()

    def set_window(self, window):
        self.window = window

    def enqueue_log(self, message: str):
        if message:
            self.log_queue.put(str(message))

    def run_automation(self):
        with self.state_lock:
            if self.is_running:
                return {"success": False, "error": "Ya hay una ejecución en curso."}
            self.is_running = True
            self.final_status = None
            self.cancel_event.clear()

        self.enqueue_log("Iniciando proceso principal...")
        thread = threading.Thread(target=self._execute_process, name="automation-worker", daemon=True)
        self.worker_thread = thread
        thread.start()
        return {"success": True}

    def _execute_process(self):
        try:
            process_all_tickets(abort_event=self.cancel_event)
            if self.cancel_event.is_set():
                self.final_status = {
                    "success": False,
                    "message": "La ejecución fue cancelada durante el cierre de la aplicación.",
                }
            else:
                self.final_status = {
                    "success": True,
                    "message": "Proceso finalizado exitosamente.",
                }
        except Exception as e:
            get_logger("GUI").error("Error fatal en proceso: %s", e)
            self.final_status = {"success": False, "message": str(e)}
        finally:
            with self.state_lock:
                self.is_running = False

    def get_runtime_updates(self):
        logs: list[str] = []
        while True:
            try:
                logs.append(self.log_queue.get_nowait())
            except queue.Empty:
                break

        return {
            "logs": logs,
            "is_running": self.is_running,
            "final_status": self.final_status,
        }

    def request_stop(self):
        self.cancel_event.set()
        self.enqueue_log("Se solicitó detener el proceso actual.")
        return {"success": True}

    def handle_window_closing(self):
        if self.shutdown_started.is_set():
            return

        self.shutdown_started.set()
        self.cancel_event.set()
        self.enqueue_log("Cerrando aplicación. Se solicitó cancelar la ejecución activa.")

        shutdown_thread = threading.Thread(
            target=self._finish_shutdown,
            name="gui-shutdown",
            daemon=True,
        )
        shutdown_thread.start()

    def _finish_shutdown(self):
        worker = self.worker_thread
        if worker and worker.is_alive():
            worker.join(timeout=8)

        # Si algo quedó colgado, soltamos el servicio del driver y cerramos el proceso.
        kill_processes(["msedgedriver.exe"])
        time.sleep(1)
        os._exit(0)

    def get_config(self):
        try:
            return get_settings().to_dict()
        except Exception as e:
            return {"error": str(e)}

    def save_config(self, data):
        try:
            self.settings_manager.save_settings(data)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}


def start_gui():
    api = ProcessorAPI()
    html_path = get_resource_path(os.path.join("src", "gui", "index.html"))

    window = webview.create_window(
        "EssilorLuxottica - HDA Automation",
        html_path,
        js_api=api,
        width=1200,
        height=800,
        background_color="#122033",
        confirm_close=True,
    )

    api.set_window(window)
    add_gui_callback(api.enqueue_log)
    window.events.closing += api.handle_window_closing
    webview.start(debug=False)


if __name__ == "__main__":
    start_gui()
