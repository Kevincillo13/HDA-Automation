import os
from dataclasses import dataclass

from dotenv import load_dotenv


load_dotenv(os.path.join("config", ".env.local"))
load_dotenv(os.path.join("config", ".env.example"))


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


def get_settings() -> Settings:
    return Settings()
