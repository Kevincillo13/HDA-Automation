import os
from dataclasses import dataclass

from dotenv import load_dotenv


import sys
from pathlib import Path

def get_resource_path(relative_path):
    """Obtiene la ruta absoluta para recursos, compatible con PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path.cwd() / relative_path

# Intentar cargar .env.local de forma robusta
env_path = get_resource_path(os.path.join("config", ".env.local"))
if env_path.exists():
    load_dotenv(str(env_path))
else:
    # Fallback al .env estándar
    alt_env_path = get_resource_path(os.path.join("config", ".env"))
    if alt_env_path.exists():
        load_dotenv(str(alt_env_path))


from src.common.settings_manager import SettingsManager


def _get_int_env(name: str, default: int) -> int:
    raw_value = os.getenv(name, "").strip()
    if not raw_value:
        return default
    return int(raw_value)


@dataclass
class Settings:
    hda_url: str = os.getenv("HDA_URL", "")
    hda_username: str = os.getenv("HDA_USERNAME", "")
    hda_password: str = os.getenv("HDA_PASSWORD", "")
    browser_headless: bool = os.getenv("BROWSER_HEADLESS", "false").lower() == "true"
    browser_keep_open: bool = os.getenv("BROWSER_KEEP_OPEN", "true").lower() == "true"
    browser_slow_mo_ms: int = _get_int_env("BROWSER_SLOW_MO_MS", 0)
    browser_window_width: int = _get_int_env("BROWSER_WINDOW_WIDTH", 1600)
    browser_window_height: int = _get_int_env("BROWSER_WINDOW_HEIGHT", 1200)
    browser_binary_path: str = os.getenv(
        "BROWSER_BINARY_PATH",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    )
    evidence_enabled: bool = os.getenv("EVIDENCE_ENABLED", "false").lower() == "true"
    timezone: str = os.getenv("TIMEZONE", "America/Mexico_City")
    download_dir: str = os.getenv("DOWNLOAD_DIR", "runtime/downloads")
    output_dir: str = os.getenv("OUTPUT_DIR", "runtime/outputs")
    log_dir: str = os.getenv("LOG_DIR", "runtime/logs")
    evidence_dir: str = os.getenv("EVIDENCE_DIR", "runtime/evidence")
    template_path: str = os.getenv("TEMPLATE_PATH", "templates/template.xlsx")
    mail_test_recipient: str = os.getenv("MAIL_TEST_RECIPIENT", "")
    mail_primary_recipient: str = os.getenv("MAIL_PRIMARY_RECIPIENT", "")
    mail_secondary_recipient: str = os.getenv("MAIL_SECONDARY_RECIPIENT", "")
    mail_fms_recipient: str = os.getenv("MAIL_FMS_RECIPIENT", "")
    mail_afs_recipient: str = os.getenv("MAIL_AFS_RECIPIENT", "")
    mail_usd_recipient: str = os.getenv("MAIL_USD_RECIPIENT", "")
    mail_cad_recipient: str = os.getenv("MAIL_CAD_RECIPIENT", "")
    mail_summary_recipient: str = os.getenv("MAIL_SUMMARY_RECIPIENT", "")
    mail_error_recipient: str = os.getenv("MAIL_ERROR_RECIPIENT", "")
    mail_bcc_recipient: str = os.getenv("MAIL_BCC_RECIPIENT", "")
    mail_subject_prefix: str = os.getenv("MAIL_SUBJECT_PREFIX", "[TEST]")
    smtp_host: str = os.getenv("SMTP_HOST", "")
    smtp_port: int = _get_int_env("SMTP_PORT", 0)
    smtp_username: str = os.getenv("SMTP_USERNAME", "")
    smtp_password: str = os.getenv("SMTP_PASSWORD", "")
    smtp_sender: str = os.getenv("SMTP_SENDER", "")
    smtp_use_tls: bool = os.getenv("SMTP_USE_TLS", "true").lower() == "true"
    smtp_use_ssl: bool = os.getenv("SMTP_USE_SSL", "false").lower() == "true"
    sap_username_fms: str = os.getenv("SAP_USERNAME_FMS", "")
    sap_password_fms: str = os.getenv("SAP_PASSWORD_FMS", "")
    sap_username_afs: str = os.getenv("SAP_USERNAME_AFS", "")
    sap_password_afs: str = os.getenv("SAP_PASSWORD_AFS", "")
    sap_connection_name_fms: str = os.getenv("SAP_CONNECTION_NAME_FMS", "")
    sap_connection_name_afs: str = os.getenv("SAP_CONNECTION_NAME_AFS", "")
    sap_client_fms: str = os.getenv("SAP_CLIENT_FMS", "500")
    sap_client_afs: str = os.getenv("SAP_CLIENT_AFS", "400")
    sap_language: str = os.getenv("SAP_LANGUAGE", "EN")
    sap_tcode_fms: str = os.getenv("SAP_TCODE_FMS", "ZFIN_AP_NONPO_LUCY4")
    sap_tcode_afs: str = os.getenv("SAP_TCODE_AFS", "ZFIN_AP_NONPO_LUCY4")
    sap_executable_path: str = os.getenv(
        "SAP_EXECUTABLE_PATH",
        r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",
    )

    def to_dict(self) -> dict:
        """Convierte los ajustes a un diccionario para la GUI."""
        from dataclasses import asdict
        return asdict(self)


def get_settings() -> Settings:
    """Retorna los ajustes cargados desde ENV y sobreescritos por el JSON de la APP."""
    settings = Settings()
    manager = SettingsManager()
    manager.update_from_env(settings)
    return settings
