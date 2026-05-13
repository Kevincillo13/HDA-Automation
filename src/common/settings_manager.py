import json
import os
import sys
from pathlib import Path
from typing import Any, Dict


def _get_settings_base_dir() -> Path:
    """Resuelve la carpeta base donde viven los archivos externos de configuración."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()

class SettingsManager:
    """Maneja la persistencia de la configuración del usuario en un archivo JSON."""
    
    def __init__(self, settings_file: str = "bot_settings.json"):
        self.settings_path = _get_settings_base_dir() / settings_file
        self.legacy_settings_path = _get_settings_base_dir() / "app_settings.json"

    def load_settings(self) -> Dict[str, Any]:
        """Carga los ajustes desde el archivo JSON si existe."""
        settings_path = self.settings_path
        if not settings_path.exists() and self.legacy_settings_path.exists():
            settings_path = self.legacy_settings_path

        if not settings_path.exists():
            return {}
        
        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error cargando settings: {e}")
            return {}

    def save_settings(self, settings_dict: Dict[str, Any]) -> bool:
        """Guarda un diccionario de ajustes en el archivo JSON."""
        try:
            # No guardamos objetos complejos, solo strings, ints y bools
            serializable_settings = {
                k: v for k, v in settings_dict.items() 
                if isinstance(v, (str, int, bool, float)) or v is None
            }
            
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(serializable_settings, f, indent=4)
            return True
        except Exception as e:
            print(f"Error guardando settings: {e}")
            return False

    def update_from_env(self, settings_obj: Any):
        """
        Toma un objeto Settings existente y lo actualiza con los valores 
        guardados en el JSON (si existen).
        """
        stored = self.load_settings()
        for key, value in stored.items():
            if hasattr(settings_obj, key):
                setattr(settings_obj, key, value)
