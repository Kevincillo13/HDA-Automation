import os
import sys
import threading
import webview
import time
from src.common.config import get_settings
from src.common.logger import get_logger, add_gui_callback
from src.common.settings_manager import SettingsManager
from src.hda_web.ticket_processing import process_all_tickets

def get_resource_path(relative_path):
    """Obtiene la ruta absoluta para recursos, compatible con PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class ProcessorAPI:
    """API expuesta a JavaScript para controlar la automatización."""
    
    def __init__(self):
        self.cancel_event = threading.Event()
        self.is_running = False
        self.window = None
        self.settings_manager = SettingsManager()

    def set_window(self, window):
        self.window = window

    def run_automation(self):
        if self.is_running:
            return {"success": False, "error": "Ya hay una ejecución en curso."}
        self.is_running = True
        self.cancel_event.clear()
        thread = threading.Thread(target=self._execute_process)
        thread.daemon = True
        thread.start()
        return {"success": True}

    def _execute_process(self):
        try:
            process_all_tickets(abort_event=self.cancel_event)
        except Exception as e:
            get_logger("GUI").error(f"Error fatal en proceso: {e}")
        finally:
            self.is_running = False

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
    
    # Regresamos a ventana ESTÁNDAR
    window = webview.create_window(
        'EssilorLuxottica - HDA Automation', 
        html_path, 
        js_api=api,
        width=1200,
        height=800,
        background_color='#122033'
    )
    
    api.set_window(window)

    def log_to_gui(message):
        if api.window:
            safe_msg = message.replace('"', '\\"').replace("'", "\\'")
            try:
                api.window.evaluate_js(f"addLog('{safe_msg}')")
            except:
                pass
        
    add_gui_callback(log_to_gui)
    webview.start(debug=False)

if __name__ == "__main__":
    start_gui()
