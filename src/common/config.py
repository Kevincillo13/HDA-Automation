import os
from dataclasses import dataclass

from dotenv import load_dotenv


load_dotenv(os.path.join("config", ".env.local"))
load_dotenv(os.path.join("config", ".env.example"))


@dataclass
class Settings:
    hda_url: str = os.getenv("HDA_URL", "")
    hda_username: str = os.getenv("HDA_USERNAME", "")
    hda_password: str = os.getenv("HDA_PASSWORD", "")
    browser_headless: bool = os.getenv("BROWSER_HEADLESS", "false").lower() == "true"
    browser_keep_open: bool = os.getenv("BROWSER_KEEP_OPEN", "true").lower() == "true"
    browser_slow_mo_ms: int = int(os.getenv("BROWSER_SLOW_MO_MS", "0"))
    browser_window_width: int = int(os.getenv("BROWSER_WINDOW_WIDTH", "1600"))
    browser_window_height: int = int(os.getenv("BROWSER_WINDOW_HEIGHT", "1200"))
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


def get_settings() -> Settings:
    return Settings()
