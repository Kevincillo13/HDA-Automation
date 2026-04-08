from __future__ import annotations

import subprocess
import time
from typing import Any

from src.common.config import Settings, get_settings
from src.common.logger import get_logger

try:
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - depends on Windows environment
    win32com = None  # type: ignore[assignment]


class SAPGuiClient:
    """Base client for SAP GUI Scripting login and session access."""

    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()
        self.logger = get_logger("sap_client")
        self.application: Any | None = None
        self.connection: Any | None = None
        self.session: Any | None = None

    def start(self) -> None:
        """Launch SAP Logon and wait until the scripting engine is available."""
        self._require_pywin32()

        sap_path = self.settings.sap_executable_path.strip()
        if not sap_path:
            raise ValueError("SAP executable path is not configured.")

        self.logger.info("Starting SAP Logon: %s", sap_path)
        subprocess.Popen(sap_path)
        self._wait_for_sap_gui()

    def login(self) -> Any:
        """Open the configured connection and perform SAP login."""
        self._require_pywin32()

        if not self.settings.sap_connection_name.strip():
            raise ValueError("SAP connection name is not configured.")
        if not self.settings.sap_username.strip() or not self.settings.sap_password.strip():
            raise ValueError("SAP credentials are not configured.")

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        self.application = sap_gui_auto.GetScriptingEngine
        self.logger.info("Opening SAP connection: %s", self.settings.sap_connection_name)
        self.connection = self.application.OpenConnection(
            self.settings.sap_connection_name,
            True,
        )
        self.session = self._wait_for_session(self.connection)
        self.session.findById("wnd[0]").maximize()
        self.logger.info("SAP session opened.")

        self._fill_login_form()
        self._handle_multiple_logon_popup()
        self.session = self._refresh_active_session()
        self.logger.info("SAP session ready for automation.")
        return self.session

    def close(self) -> None:
        """Release SAP references kept by the client."""
        self.session = None
        self.connection = None
        self.application = None

    def open_tcode(self, tcode: str) -> None:
        """Open a transaction code, preferring SAP native transaction navigation."""
        session = self.get_session()
        normalized_tcode = tcode.strip()
        if not normalized_tcode:
            raise ValueError("TCode is empty.")

        self.logger.info("Opening SAP TCode: %s", normalized_tcode)
        sanitized_tcode = normalized_tcode
        if sanitized_tcode.lower().startswith("/n"):
            sanitized_tcode = sanitized_tcode[2:]
        sanitized_tcode = sanitized_tcode.lstrip("/")

        try:
            session.StartTransaction(Transaction=sanitized_tcode)
            self.logger.info("SAP TCode opened via StartTransaction: %s", sanitized_tcode)
            return
        except Exception as exc:
            self.logger.warning(
                "StartTransaction failed for %s, falling back to command field: %s",
                sanitized_tcode,
                exc,
            )

        command = f"/n{sanitized_tcode}"
        try:
            command_field = session.findById("wnd[0]/tbar[0]/okcd")
        except Exception as exc:
            raise RuntimeError(
                "SAP command field 'wnd[0]/tbar[0]/okcd' was not found."
            ) from exc

        try:
            command_field.setFocus()
        except Exception:
            pass

        try:
            command_field.text = ""
        except Exception:
            pass

        try:
            command_field.caretPosition = 0
        except Exception:
            pass

        command_field.text = command

        try:
            command_field.caretPosition = len(command)
        except Exception:
            pass

        field_value = ""
        try:
            field_value = str(command_field.text)
        except Exception:
            field_value = "<unreadable>"

        self.logger.info(
            "SAP command field populated | requested=%s | actual=%s",
            command,
            field_value,
        )

        try:
            session.findById("wnd[0]/tbar[0]/btn[0]").press()
        except Exception:
            session.sendVKey(0)
        self.logger.info("SAP TCode submitted: %s", command)

    def get_session(self) -> Any:
        if not self.session:
            raise RuntimeError("SAP session is not initialized.")
        return self.session

    def _require_pywin32(self) -> None:
        if win32com is None:
            raise RuntimeError(
                "pywin32 is not installed. Install it before using SAP GUI automation."
            )

    def _wait_for_sap_gui(self, timeout_seconds: int = 30) -> None:
        for attempt in range(timeout_seconds):
            try:
                win32com.client.GetObject("SAPGUI")
                self.logger.info(
                    "SAP GUI scripting engine detected after %s second(s).",
                    attempt + 1,
                )
                return
            except Exception:
                time.sleep(1)
        raise RuntimeError("SAP GUI scripting engine was not detected in time.")

    def _wait_for_session(self, connection: Any, timeout_seconds: int = 30) -> Any:
        for _ in range(timeout_seconds):
            if connection.Children.Count > 0:
                return connection.Children(0)
            time.sleep(1)
        raise RuntimeError("No SAP session became available after opening the connection.")

    def _fill_login_form(self) -> None:
        session = self.get_session()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.settings.sap_username
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.settings.sap_password
        session.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.settings.sap_client
        session.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.settings.sap_language
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        self.logger.info("SAP login submitted.")

    def _handle_multiple_logon_popup(self) -> None:
        session = self.get_session()
        time.sleep(2)
        try:
            popup = session.findById("wnd[1]")
            popup.findById("usr/radMULTI_LOGON_OPT2").select()
            popup.findById("tbar[0]/btn[0]").press()
            self.logger.info("Handled SAP multiple logon popup.")
        except Exception:
            self.logger.info("No SAP multiple logon popup was detected.")

    def _refresh_active_session(self) -> Any:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        if application.Children.Count == 0:
            raise RuntimeError("No SAP connections are available after login.")

        connection = application.Children(0)
        if connection.Children.Count == 0:
            raise RuntimeError("No SAP sessions are available after login.")

        self.application = application
        self.connection = connection
        return connection.Children(0)


def test_sap_login() -> None:
    """Standalone helper to verify SAP login flow."""
    client = SAPGuiClient()
    client.start()
    client.login()
    print("SAP session ready.")


def test_sap_tcode(tcode: str) -> None:
    """Standalone helper to verify TCode navigation after login."""
    client = SAPGuiClient()
    client.start()
    client.login()
    client.open_tcode(tcode)
    print(f"SAP TCode opened: {tcode}")


if __name__ == "__main__":
    test_sap_tcode("ZFIN_AP_NONPO_LUCY4")
