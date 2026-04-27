from __future__ import annotations

import csv
import subprocess
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from src.common.config import Settings, get_settings
from src.common.logger import get_logger

try:
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - depends on Windows environment
    win32com = None  # type: ignore[assignment]


class SAPGuiClient:
    """Base client for SAP GUI Scripting login and session access."""

    CATEGORY_ACTIONS = {
        "success": "accept",
        "company_code_invalid": "suspend_row",
        "one_time_vendor_invalid": "suspend_row",
        "vendor_invalid": "suspend_row",
        "cost_center_invalid": "suspend_row",
        "profit_center_invalid": "suspend_row",
        "profit_center_missing": "suspend_row",
        "gl_account_invalid": "suspend_row",
        "duplicate_document": "suspend_row",
        "invoice_date_invalid": "suspend_row",
        "invoice_date_in_future": "suspend_row",
        "gl_account_missing": "fix_local_data",
        "amount_invalid": "fix_csv_format",
        "company_code_missing": "fix_local_data",
        "empty_message": "manual_review",
        "generic_document_error": "manual_review",
        "unknown": "manual_review",
    }

    CATEGORY_PRIORITY = {
        "company_code_invalid": 100,
        "one_time_vendor_invalid": 95,
        "vendor_invalid": 90,
        "gl_account_missing": 89,
        "gl_account_invalid": 88,
        "cost_center_invalid": 87,
        "profit_center_invalid": 86,
        "profit_center_missing": 85,
        "amount_invalid": 84,
        "invoice_date_invalid": 83,
        "invoice_date_in_future": 82,
        "duplicate_document": 80,
        "generic_document_error": 50,
        "unknown": 40,
        "empty_message": 20,
        "success": 0,
    }

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

    def login(self, system_group: str) -> Any:
        """Open the configured connection for the given system group and perform SAP login."""
        self._require_pywin32()

        connection_name = self._resolve_connection_name(system_group)
        username, password = self._resolve_credentials(system_group)
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        self.application = sap_gui_auto.GetScriptingEngine
        self.logger.info(
            "Opening SAP connection | system_group=%s | connection=%s",
            system_group,
            connection_name,
        )
        self.connection = self.application.OpenConnection(
            connection_name,
            True,
        )
        self.session = self._wait_for_session(self.connection)
        self.session.findById("wnd[0]").maximize()
        self.logger.info("SAP session opened.")

        client_number = self._resolve_client_number(system_group)
        self._fill_login_form(username, password, client_number)
        self._handle_multiple_logon_popup()
        self.session = self._refresh_active_session()
        self.logger.info("SAP session ready for automation.")
        return self.session

    def open_validation_context(self, csv_path: str) -> tuple[str, str]:
        """Login to the correct SAP system and open the validation TCode for a CSV."""
        system_group = self.infer_system_group_from_csv_path(csv_path)
        tcode = self._resolve_tcode(system_group)
        self.start()
        self.login(system_group)
        self.open_tcode(tcode)
        self.logger.info(
            "SAP validation context ready | csv=%s | system_group=%s | tcode=%s",
            Path(csv_path).name,
            system_group,
            tcode,
        )
        return system_group, tcode

    def close(self) -> None:
        """Release SAP references and attempt to close the connection."""
        try:
            if self.connection:
                self.logger.info("Closing SAP connection.")
                # Attempt to close the connection to release SAP GUI windows
                self.connection.CloseConnection()
        except Exception as exc:
            self.logger.warning("Failed to close SAP connection: %s", exc)
        
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

    def fill_validation_form(
        self,
        csv_path: str,
        posting_date: str | None = None,
        company_code: str | None = None,
        currency: str | None = None,
        separator: str = ",",
        test_mode: bool = True,
    ) -> dict[str, str]:
        """Fill the SAP validation screen using the IDs discovered in the session."""
        session = self.get_session()
        csv_metadata = self._infer_csv_metadata(csv_path)

        resolved_company_code = (company_code or csv_metadata["company_code"]).strip()
        resolved_currency = (currency or csv_metadata["currency"]).strip().upper()
        resolved_posting_date = posting_date or datetime.now().strftime("%m%d%Y")
        resolved_separator = separator or ","

        if not resolved_company_code:
            raise ValueError("Company code could not be inferred for SAP validation.")
        if not resolved_currency:
            raise ValueError("Currency could not be inferred for SAP validation.")

        absolute_csv_path = str(Path(csv_path).resolve())
        self.logger.info(
            "Filling SAP validation form | file=%s | posting_date=%s | company_code=%s | currency=%s | separator=%s | test_mode=%s",
            absolute_csv_path,
            resolved_posting_date,
            resolved_company_code,
            resolved_currency,
            resolved_separator,
            test_mode,
        )

        time.sleep(1)
        self._set_form_field_text("wnd[0]/usr/ctxtP_FILE", absolute_csv_path, "P_FILE")
        self._set_form_field_text("wnd[0]/usr/ctxtP_DATUM", resolved_posting_date, "P_DATUM")
        self._set_form_field_text("wnd[0]/usr/ctxtP_BUKRS", resolved_company_code, "P_BUKRS")
        self._set_form_field_text("wnd[0]/usr/txtP_WAERS", resolved_currency, "P_WAERS")
        self._set_form_field_text("wnd[0]/usr/txtP_FSEP", resolved_separator, "P_FSEP")
        self._set_checkbox_value("wnd[0]/usr/chkP_TEST", bool(test_mode), "P_TEST")

        return {
            "file": absolute_csv_path,
            "posting_date": resolved_posting_date,
            "company_code": resolved_company_code,
            "currency": resolved_currency,
            "separator": resolved_separator,
            "test_mode": str(bool(test_mode)),
        }

    def execute_validation(self) -> None:
        """Execute SAP validation from the current validation screen using F8."""
        session = self.get_session()
        self.logger.info("Executing SAP validation with F8.")
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

    def navigate_back_to_validation_form(self) -> None:
        """Return from the SAP results screen to the validation form using F3."""
        session = self.get_session()
        self.logger.info("Preparing to return to SAP validation form. Waiting 3 seconds before navigating back.")
        time.sleep(3)

        for attempt in range(1, 4):
            self.logger.info("Returning to SAP validation form with toolbar Back button | attempt=%s", attempt)
            self._log_current_sap_state()

            back_used = False
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
                back_used = True
                self.logger.info("SAP toolbar Back button pressed successfully | attempt=%s", attempt)
            except Exception as exc:
                self.logger.warning(
                    "SAP toolbar Back button failed | attempt=%s | error=%s. Falling back to F3.",
                    attempt,
                    exc,
                )
                try:
                    session.sendVKey(3)
                    self.logger.info("SAP F3 fallback sent successfully | attempt=%s", attempt)
                except Exception as fallback_exc:
                    self.logger.warning(
                        "SAP F3 fallback also failed | attempt=%s | error=%s",
                        attempt,
                        fallback_exc,
                    )

            if back_used:
                time.sleep(2)
            else:
                time.sleep(2)

            if self._wait_for_validation_form(timeout_seconds=6, raise_on_timeout=False):
                return

            self.logger.warning(
                "SAP validation form not detected after navigation attempt %s. Collecting SAP state and trying fallback.",
                attempt,
            )
            self._log_current_sap_state()
            self._handle_optional_popup()

            if self._wait_for_validation_form(timeout_seconds=3, raise_on_timeout=False):
                return

            if self._is_results_grid_visible():
                self.logger.warning(
                    "SAP results grid still visible after F3 attempt %s. Trying toolbar Back button.",
                    attempt,
                )
                try:
                    session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    time.sleep(2)
                except Exception as exc:
                    self.logger.warning(
                        "SAP toolbar Back button failed during fallback | attempt=%s | error=%s",
                        attempt,
                        exc,
                    )

                if self._wait_for_validation_form(timeout_seconds=4, raise_on_timeout=False):
                    return

        raise RuntimeError("SAP validation form was not detected after multiple F3 attempts.")

    def inspect_validation_results(
        self,
        grid_id: str = "wnd[0]/shellcont/shell/shellcont[1]/shell",
        sample_rows: int | None = None,
    ) -> dict[str, Any]:
        """Read the SAP validation result grid and log its structure for reverse engineering."""
        session = self.get_session()
        self.logger.info("Inspecting SAP validation results grid: %s", grid_id)

        grid = session.findById(grid_id)
        row_count = self._safe_int(self._sap_attr(grid, "RowCount"), default=0)
        column_count = self._safe_int(self._sap_attr(grid, "ColumnCount"), default=0)
        column_order = self._to_list(self._sap_attr(grid, "ColumnOrder"))
        current_cell_column = self._sap_attr(grid, "currentCellColumn")
        current_cell_row = self._sap_attr(grid, "currentCellRow")

        columns: list[dict[str, str]] = []
        for column_key in column_order:
            title = self._get_column_title(grid, column_key)
            columns.append({"key": column_key, "title": title})

        if not columns and column_count:
            for index in range(column_count):
                column_key = f"index_{index}"
                columns.append({"key": column_key, "title": ""})

        sample_data: list[dict[str, str]] = []
        effective_rows = max(row_count, 0) if sample_rows is None else min(max(row_count, 0), sample_rows)
        for row_index in range(effective_rows):
            row_payload: dict[str, str] = {"__row__": str(row_index)}
            for column in columns:
                column_key = column["key"]
                if not column_key or column_key.startswith("index_"):
                    continue
                row_payload[column_key] = self._get_cell_value(grid, row_index, column_key)
            sample_data.append(row_payload)

        self.logger.info(
            "SAP validation grid summary | grid_id=%s | row_count=%s | column_count=%s | current_cell_row=%s | current_cell_column=%s",
            grid_id,
            row_count,
            column_count,
            current_cell_row,
            current_cell_column,
        )
        for column in columns:
            self.logger.info(
                "SAP validation grid column | key=%s | title=%s",
                column["key"],
                column["title"] or "<empty>",
            )
        classified_rows = [self._classify_validation_row(row_payload) for row_payload in sample_data]
        grouped_by_invoice = self._group_validation_rows_by_invoice(classified_rows)

        for row_payload in sample_data:
            self.logger.info("SAP validation grid row | %s", row_payload)
        for classified_row in classified_rows:
            self.logger.info(
                "SAP validation classified row | invoice=%s | category=%s | success=%s | error=%s | message=%s",
                classified_row.get("invoice_number", ""),
                classified_row.get("category", ""),
                classified_row.get("is_success"),
                classified_row.get("is_error"),
                classified_row.get("message", ""),
            )
        for invoice_number, invoice_summary in grouped_by_invoice.items():
            self.logger.info(
                "SAP validation invoice summary | invoice=%s | categories=%s | primary_category=%s | action=%s | has_error=%s | success_messages=%s | error_messages=%s",
                invoice_number,
                invoice_summary.get("categories", []),
                invoice_summary.get("primary_category", ""),
                invoice_summary.get("action", ""),
                invoice_summary.get("has_error"),
                len(invoice_summary.get("success_messages", [])),
                len(invoice_summary.get("error_messages", [])),
            )

        if row_count == 0:
            self.logger.warning(
                "SAP validation returned row_count=0. This usually means format/export/parse issue, not a clean business validation."
            )

        return {
            "grid_id": grid_id,
            "row_count": row_count,
            "column_count": column_count,
            "columns": columns,
            "sample_rows": sample_data,
            "classified_rows": classified_rows,
            "grouped_by_invoice": grouped_by_invoice,
            "decision_summary": self._build_decision_summary(grouped_by_invoice, row_count),
        }

    def prepare_and_execute_validation(
        self,
        csv_path: str,
        posting_date: str | None = None,
        company_code: str | None = None,
        currency: str | None = None,
        separator: str = ",",
        test_mode: bool = True,
    ) -> dict[str, Any]:
        """Open SAP context for a CSV, fill the validation form and run it."""
        self._log_csv_payload(csv_path)
        system_group, tcode = self.open_validation_context(csv_path)
        payload = self.fill_validation_form(
            csv_path=csv_path,
            posting_date=posting_date,
            company_code=company_code,
            currency=currency,
            separator=separator,
            test_mode=test_mode,
        )
        self.execute_validation()
        validation_results = self.inspect_validation_results()
        payload["system_group"] = system_group
        payload["tcode"] = tcode
        payload["validation_results"] = validation_results
        return payload

    def validate_csv_until_clean(
        self,
        csv_path: str,
        posting_date: str | None = None,
        company_code: str | None = None,
        currency: str | None = None,
        separator: str = ",",
        test_mode: bool = True,
        max_iterations: int = 10,
        retry_suffix: str | None = None,
    ) -> dict[str, Any]:
        """Run SAP validation in a loop until all remaining rows are valid or the process gets blocked."""
        current_csv_path = str(Path(csv_path).resolve())
        iteration_history: list[dict[str, Any]] = []
        all_suspended_invoices: set[str] = set()
        suspension_reasons: dict[str, list[str]] = {}
        system_group = self.infer_system_group_from_csv_path(current_csv_path)
        tcode = self._resolve_tcode(system_group)

        for iteration in range(1, max_iterations + 1):
            self.logger.info(
                "SAP validation cycle start | iteration=%s | csv=%s",
                iteration,
                current_csv_path,
            )

            if iteration == 1:
                self._log_csv_payload(current_csv_path)
                self.open_validation_context(current_csv_path)
            else:
                self.navigate_back_to_validation_form()
                self._log_csv_payload(current_csv_path)

            payload = self.fill_validation_form(
                csv_path=current_csv_path,
                posting_date=posting_date,
                company_code=company_code,
                currency=currency,
                separator=separator,
                test_mode=test_mode,
            )
            self.execute_validation()
            validation_results = self.inspect_validation_results()
            validation_payload = {
                **payload,
                "system_group": system_group,
                "tcode": tcode,
                "validation_results": validation_results,
            }
            decision_summary = validation_results["decision_summary"]

            iteration_history.append(
                {
                    "iteration": iteration,
                    "csv_path": current_csv_path,
                    "decision_summary": decision_summary,
                }
            )

            if not decision_summary.get("has_blocking_errors", False):
                self.logger.info(
                    "SAP validation cycle clean | iteration=%s | csv=%s",
                    iteration,
                    current_csv_path,
                )
                return {
                    "status": "clean",
                    "final_csv_path": current_csv_path,
                    "iterations": iteration_history,
                    "last_validation_payload": validation_payload,
                    "all_suspended_invoices": sorted(list(all_suspended_invoices)),
                    "suspension_reasons": suspension_reasons,
                }

            # Collect all invoices that need to be removed to proceed
            blocking_invoices = (
                decision_summary.get("suspended_invoices", []) +
                decision_summary.get("fix_local_data_invoices", []) +
                decision_summary.get("fix_csv_format_invoices", []) +
                decision_summary.get("manual_review_invoices", [])
            )
            
            # Filter empty values and deduplicate
            unique_blocking = sorted(list(set(str(inv).strip() for inv in blocking_invoices if inv and str(inv).strip())))

            if unique_blocking:
                all_suspended_invoices.update(unique_blocking)
                
                # Capturar motivos de suspensión para el reporte
                grouped = validation_results.get("grouped_by_invoice", {})
                for inv in unique_blocking:
                    if inv in grouped:
                        msgs = grouped[inv].get("error_messages", [])
                        if msgs:
                            suspension_reasons[inv] = msgs

                next_csv_path = self._build_retry_csv_without_invoices(
                    current_csv_path,
                    unique_blocking,
                    iteration,
                    retry_suffix=retry_suffix,
                )
                if next_csv_path == current_csv_path:
                    self.logger.warning(
                        "SAP validation cycle could not produce a different retry CSV. Stopping as blocked."
                    )
                    return {
                        "status": "blocked",
                        "reason": "retry_csv_not_changed",
                        "final_csv_path": current_csv_path,
                        "iterations": iteration_history,
                        "last_validation_payload": validation_payload,
                        "all_suspended_invoices": sorted(list(all_suspended_invoices)),
                        "suspension_reasons": suspension_reasons,
                    }
                current_csv_path = next_csv_path
                continue

            self.logger.warning(
                "SAP validation cycle blocked by non-auto-resolvable issues with no identifiable invoices."
            )
            return {
                "status": "blocked",
                "reason": "non_auto_resolvable_errors_no_invoices",
                "final_csv_path": current_csv_path,
                "iterations": iteration_history,
                "last_validation_payload": validation_payload,
                "all_suspended_invoices": sorted(list(all_suspended_invoices)),
                "suspension_reasons": suspension_reasons,
            }

        self.logger.warning(
            "SAP validation cycle reached max iterations without becoming clean | csv=%s | max_iterations=%s",
            current_csv_path,
            max_iterations,
        )
        return {
            "status": "max_iterations_reached",
            "final_csv_path": current_csv_path,
            "iterations": iteration_history,
            "all_suspended_invoices": sorted(list(all_suspended_invoices)),
            "suspension_reasons": suspension_reasons,
        }

    def get_session(self) -> Any:
        if not self.session:
            raise RuntimeError("SAP session is not initialized.")
        return self.session

    def infer_system_group_from_csv_path(self, csv_path: str) -> str:
        csv_name = Path(csv_path).name.upper()
        if "_FMS_" in csv_name:
            return "FMS"
        if "_AFS_" in csv_name:
            return "AFS"
        raise ValueError(f"Unable to infer SAP system group from CSV name: {csv_name}")

    def _resolve_connection_name(self, system_group: str) -> str:
        normalized = system_group.strip().upper()
        if normalized == "FMS":
            connection_name = self.settings.sap_connection_name_fms.strip()
        elif normalized == "AFS":
            connection_name = self.settings.sap_connection_name_afs.strip()
        else:
            raise ValueError(f"Unsupported SAP system group: {system_group}")

        if not connection_name:
            raise ValueError(f"SAP connection name is not configured for {normalized}.")
        return connection_name

    def _resolve_tcode(self, system_group: str) -> str:
        normalized = system_group.strip().upper()
        if normalized == "FMS":
            tcode = self.settings.sap_tcode_fms.strip()
        elif normalized == "AFS":
            tcode = self.settings.sap_tcode_afs.strip()
        else:
            raise ValueError(f"Unsupported SAP system group: {system_group}")

        if not tcode:
            raise ValueError(f"SAP TCode is not configured for {normalized}.")
        return tcode

    def _resolve_credentials(self, system_group: str) -> tuple[str, str]:
        normalized = system_group.strip().upper()
        if normalized == "FMS":
            username = self.settings.sap_username_fms.strip()
            password = self.settings.sap_password_fms.strip()
        elif normalized == "AFS":
            username = self.settings.sap_username_afs.strip()
            password = self.settings.sap_password_afs.strip()
        else:
            raise ValueError(f"Unsupported SAP system group: {system_group}")

        if not username or not password:
            raise ValueError(f"SAP credentials are not configured for {normalized}.")
        return username, password

    def _resolve_client_number(self, system_group: str) -> str:
        normalized = system_group.strip().upper()
        if normalized == "FMS":
            client_number = self.settings.sap_client_fms.strip()
        elif normalized == "AFS":
            client_number = self.settings.sap_client_afs.strip()
        else:
            raise ValueError(f"Unsupported SAP system group: {system_group}")

        if not client_number:
            raise ValueError(f"SAP client number is not configured for {normalized}.")
        return client_number

    def _infer_csv_metadata(self, csv_path: str) -> dict[str, str]:
        csv_file_path = Path(csv_path)
        if not csv_file_path.exists():
            raise FileNotFoundError(f"CSV file not found: {csv_file_path}")

        with csv_file_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            reader = csv.DictReader(csv_file)
            first_row = next(reader, None)

        if not first_row:
            raise ValueError(f"CSV file has no data rows: {csv_file_path}")

        company_code = str(first_row.get("Company Code", "")).strip()
        currency = str(first_row.get("Currency USD/CAD", "")).strip().upper()
        return {
            "company_code": company_code,
            "currency": currency,
        }

    def _read_csv_rows(self, csv_path: str) -> tuple[list[str], list[dict[str, str]]]:
        csv_file_path = Path(csv_path)
        if not csv_file_path.exists():
            raise FileNotFoundError(f"CSV file not found: {csv_file_path}")

        with csv_file_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            reader = csv.DictReader(csv_file)
            fieldnames = [field.strip() for field in (reader.fieldnames or []) if field and field.strip()]
            rows = [{str(key).strip(): str(value).strip() for key, value in row.items()} for row in reader]
        return fieldnames, rows

    def _log_csv_payload(self, csv_path: str) -> None:
        fieldnames, rows = self._read_csv_rows(csv_path)
        self.logger.info(
            "SAP validation CSV snapshot | file=%s | rows=%s | columns=%s",
            Path(csv_path).resolve(),
            len(rows),
            fieldnames,
        )
        for index, row in enumerate(rows, start=1):
            self.logger.info("SAP validation CSV row | row_number=%s | %s", index, row)

    def _classify_validation_row(self, row_payload: dict[str, str]) -> dict[str, Any]:
        message = str(row_payload.get("MESSAGES", "")).strip()
        normalized_message = message.lower()

        category = "unknown"
        is_success = False
        is_error = True

        if not message:
            category = "empty_message"
        elif "document can be created with company code" in normalized_message:
            category = "success"
            is_success = True
            is_error = False
        elif "check whether document has already been entered" in normalized_message:
            category = "duplicate_document"
        elif "company code not exist" in normalized_message:
            category = "company_code_invalid"
        elif "invalid one time vendor number" in normalized_message:
            category = "one_time_vendor_invalid"
        elif "vendor " in normalized_message and " is not defined in company code " in normalized_message:
            category = "vendor_invalid"
        elif "cost center " in normalized_message and " does not exist " in normalized_message:
            category = "cost_center_invalid"
        elif "profit center" in normalized_message and "not filled" in normalized_message:
            category = "profit_center_missing"
        elif "profit center" in normalized_message and (
            "does not exist" in normalized_message or "not found" in normalized_message
        ):
            category = "profit_center_invalid"
        elif "required field gl_account" in normalized_message:
            category = "gl_account_missing"
        elif "company code is missing" in normalized_message:
            category = "company_code_missing"
        elif "g/l account " in normalized_message and " is not defined in chart of accounts " in normalized_message:
            category = "gl_account_invalid"
        elif "incorrect value in amount field" in normalized_message:
            category = "amount_invalid"
        elif "date " in normalized_message and " is invalid; use the format " in normalized_message:
            category = "invoice_date_invalid"
        elif " cannot exceed " in normalized_message:
            category = "invoice_date_in_future"
        elif "error in document:" in normalized_message:
            category = "generic_document_error"

        return {
            "row_index": row_payload.get("__row__", ""),
            "line_no": str(row_payload.get("LINE_NO", "")).strip(),
            "invoice_number": str(row_payload.get("XBLNR", "")).strip() or "<blank>",
            "vendor": str(row_payload.get("LIFNR", "")).strip(),
            "company_message": message,
            "message": message,
            "category": category,
            "is_success": is_success,
            "is_error": is_error,
            "raw": row_payload,
        }

    def _group_validation_rows_by_invoice(
        self,
        classified_rows: list[dict[str, Any]],
    ) -> dict[str, dict[str, Any]]:
        grouped: dict[str, dict[str, Any]] = {}
        grouped_messages: dict[str, list[dict[str, Any]]] = defaultdict(list)

        for row in classified_rows:
            grouped_messages[row["invoice_number"]].append(row)

        for invoice_number, rows in grouped_messages.items():
            success_messages = [row["message"] for row in rows if row["is_success"]]
            error_messages = [row["message"] for row in rows if row["is_error"]]
            categories = sorted({str(row["category"]) for row in rows if row.get("category")})
            primary_error_row = self._select_primary_error_row(rows)
            primary_category = (
                str(primary_error_row.get("category"))
                if primary_error_row
                else ("success" if success_messages else "unknown")
            )
            action = self._resolve_invoice_action(primary_category, rows)
            grouped[invoice_number] = {
                "invoice_number": invoice_number,
                "line_numbers": sorted({str(row.get("line_no", "")) for row in rows if str(row.get("line_no", ""))}),
                "categories": categories,
                "has_error": any(row["is_error"] for row in rows),
                "has_success": any(row["is_success"] for row in rows),
                "success_messages": success_messages,
                "error_messages": error_messages,
                "primary_category": primary_category,
                "primary_message": (
                    str(primary_error_row.get("message"))
                    if primary_error_row
                    else (success_messages[0] if success_messages else "")
                ),
                "action": action,
                "rows": rows,
            }

        return grouped

    def _select_primary_error_row(self, rows: list[dict[str, Any]]) -> dict[str, Any] | None:
        error_rows = [row for row in rows if row.get("is_error")]
        if not error_rows:
            return None
        return max(
            error_rows,
            key=lambda row: self.CATEGORY_PRIORITY.get(str(row.get("category", "unknown")), 0),
        )

    def _resolve_invoice_action(
        self,
        primary_category: str,
        rows: list[dict[str, Any]],
    ) -> str:
        normalized_category = str(primary_category or "unknown")
        action = self.CATEGORY_ACTIONS.get(normalized_category, "manual_review")

        if normalized_category == "generic_document_error":
            specific_error_rows = [
                row
                for row in rows
                if row.get("is_error") and str(row.get("category")) != "generic_document_error"
            ]
            if specific_error_rows:
                best_specific = self._select_primary_error_row(specific_error_rows)
                if best_specific:
                    return self.CATEGORY_ACTIONS.get(
                        str(best_specific.get("category", "unknown")),
                        "manual_review",
                    )

        return action

    def _build_decision_summary(
        self,
        grouped_by_invoice: dict[str, dict[str, Any]],
        row_count: int,
    ) -> dict[str, Any]:
        accepted_invoices: list[str] = []
        suspended_invoices: list[str] = []
        fix_local_data_invoices: list[str] = []
        fix_csv_format_invoices: list[str] = []
        manual_review_invoices: list[str] = []

        if row_count == 0:
            decision_summary = {
                "accepted_invoices": [],
                "suspended_invoices": [],
                "fix_local_data_invoices": [],
                "fix_csv_format_invoices": ["<sap_empty_result_grid>"],
                "manual_review_invoices": [],
                "has_blocking_errors": True,
                "result_state": "empty_result_grid",
            }
            self.logger.info(
                "SAP validation decision summary | accepted=%s | suspended=%s | fix_local=%s | fix_csv=%s | manual_review=%s",
                decision_summary["accepted_invoices"],
                decision_summary["suspended_invoices"],
                decision_summary["fix_local_data_invoices"],
                decision_summary["fix_csv_format_invoices"],
                decision_summary["manual_review_invoices"],
            )
            return decision_summary

        for invoice_number, summary in grouped_by_invoice.items():
            action = str(summary.get("action", "manual_review"))
            if action == "accept":
                accepted_invoices.append(invoice_number)
            elif action == "suspend_row":
                suspended_invoices.append(invoice_number)
            elif action == "fix_local_data":
                fix_local_data_invoices.append(invoice_number)
            elif action == "fix_csv_format":
                fix_csv_format_invoices.append(invoice_number)
            else:
                manual_review_invoices.append(invoice_number)

        decision_summary = {
            "accepted_invoices": accepted_invoices,
            "suspended_invoices": suspended_invoices,
            "fix_local_data_invoices": fix_local_data_invoices,
            "fix_csv_format_invoices": fix_csv_format_invoices,
            "manual_review_invoices": manual_review_invoices,
            "has_blocking_errors": bool(
                suspended_invoices or fix_local_data_invoices or fix_csv_format_invoices or manual_review_invoices
            ),
            "result_state": "classified_rows",
        }

        self.logger.info(
            "SAP validation decision summary | accepted=%s | suspended=%s | fix_local=%s | fix_csv=%s | manual_review=%s",
            accepted_invoices,
            suspended_invoices,
            fix_local_data_invoices,
            fix_csv_format_invoices,
            manual_review_invoices,
        )
        return decision_summary

    def _build_retry_csv_without_invoices(
        self,
        csv_path: str,
        suspended_invoices: list[str],
        iteration: int,
        retry_suffix: str | None = None,
    ) -> str:
        suspended_invoice_set = {
            str(invoice).strip() for invoice in suspended_invoices if str(invoice).strip()
        }
        fieldnames, rows = self._read_csv_rows(csv_path)
        kept_rows = [
            row
            for row in rows
            if str(row.get("Invoice Number", "")).strip() not in suspended_invoice_set
        ]

        if not kept_rows:
            self.logger.warning(
                "SAP retry CSV generation would result in 0 rows. Skipping file creation | csv=%s",
                csv_path
            )
            # Devolvemos la ruta original; el loop detectará que no hay cambios y se detendrá
            return str(Path(csv_path).resolve())

        if len(kept_rows) == len(rows):
            self.logger.warning(
                "SAP retry CSV generation removed 0 rows | csv=%s | suspended_invoices=%s",
                csv_path,
                sorted(suspended_invoice_set),
            )
            return str(Path(csv_path).resolve())

        source_path = Path(csv_path).resolve()
        
        # Build unique filename using suffix if provided
        base_name = retry_suffix if retry_suffix else "sapr"
        temp_name = f"{base_name}_retry_{iteration:02d}{source_path.suffix}"
        next_csv_path = source_path.with_name(temp_name)
        self._write_csv_rows(next_csv_path, fieldnames, kept_rows)
        self.logger.info(
            "SAP retry CSV generated | source=%s | target=%s | removed_invoices=%s | kept_rows=%s",
            source_path,
            next_csv_path,
            sorted(suspended_invoice_set),
            len(kept_rows),
        )
        return str(next_csv_path)

    def _write_csv_rows(
        self,
        target_path: Path,
        fieldnames: list[str],
        rows: list[dict[str, str]],
    ) -> None:
        with target_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for row in rows:
                writer.writerow({field: row.get(field, "") for field in fieldnames})

    @staticmethod
    def _sap_attr(obj: Any, attr_name: str) -> Any:
        try:
            return getattr(obj, attr_name)
        except Exception:
            return None

    @staticmethod
    def _safe_int(value: Any, default: int = 0) -> int:
        try:
            return int(value)
        except Exception:
            return default

    @staticmethod
    def _to_list(value: Any) -> list[str]:
        if value is None:
            return []
        if isinstance(value, (list, tuple)):
            return [str(item).strip() for item in value if str(item).strip()]
        try:
            return [str(item).strip() for item in list(value) if str(item).strip()]
        except Exception:
            pass
        raw = str(value).strip()
        if not raw:
            return []
        return [raw]

    def _get_column_title(self, grid: Any, column_key: str) -> str:
        getters = [
            lambda: str(grid.GetColumnTitles(column_key)),
            lambda: str(grid.GetColumnTitle(column_key)),
        ]
        for getter in getters:
            try:
                value = getter().strip()
                if value:
                    return value
            except Exception:
                continue
        return ""

    def _get_cell_value(self, grid: Any, row_index: int, column_key: str) -> str:
        getters = [
            lambda: str(grid.GetCellValue(row_index, column_key)),
            lambda: str(grid.getCellValue(row_index, column_key)),
        ]
        for getter in getters:
            try:
                return getter().strip()
            except Exception:
                continue
        return ""

    def _set_form_field_text(
        self,
        field_id: str,
        value: str,
        field_label: str,
        attempts: int = 5,
        delay_seconds: float = 1.0,
    ) -> None:
        session = self.get_session()
        last_error: Exception | None = None

        for attempt in range(1, attempts + 1):
            try:
                field = session.findById(field_id)
                try:
                    field.setFocus()
                except Exception:
                    pass

                try:
                    field.text = ""
                except Exception:
                    pass

                field.text = value
                self.logger.info(
                    "SAP form field populated | field=%s | attempt=%s | value=%s",
                    field_label,
                    attempt,
                    value,
                )
                return
            except Exception as exc:
                last_error = exc
                field_state = self._describe_field_state(field_id)
                self.logger.warning(
                    "SAP form field set failed | field=%s | attempt=%s | value=%s | state=%s | error=%s",
                    field_label,
                    attempt,
                    value,
                    field_state,
                    exc,
                )
                time.sleep(delay_seconds)

        raise RuntimeError(
            f"SAP form field '{field_label}' could not be populated after {attempts} attempts."
        ) from last_error

    def _set_checkbox_value(
        self,
        field_id: str,
        value: bool,
        field_label: str,
        attempts: int = 5,
        delay_seconds: float = 1.0,
    ) -> None:
        session = self.get_session()
        last_error: Exception | None = None

        for attempt in range(1, attempts + 1):
            try:
                field = session.findById(field_id)
                field.selected = value
                try:
                    field.setFocus()
                except Exception:
                    pass
                self.logger.info(
                    "SAP checkbox populated | field=%s | attempt=%s | value=%s",
                    field_label,
                    attempt,
                    value,
                )
                return
            except Exception as exc:
                last_error = exc
                field_state = self._describe_field_state(field_id)
                self.logger.warning(
                    "SAP checkbox set failed | field=%s | attempt=%s | value=%s | state=%s | error=%s",
                    field_label,
                    attempt,
                    value,
                    field_state,
                    exc,
                )
                time.sleep(delay_seconds)

        raise RuntimeError(
            f"SAP checkbox '{field_label}' could not be populated after {attempts} attempts."
        ) from last_error

    def _describe_field_state(self, field_id: str) -> dict[str, str]:
        session = self.get_session()
        state: dict[str, str] = {"field_id": field_id}
        try:
            field = session.findById(field_id)
        except Exception as exc:
            state["exists"] = "False"
            state["lookup_error"] = str(exc)
            return state

        state["exists"] = "True"
        for attr_name in ("Type", "Name", "Text", "Changeable"):
            value = self._sap_attr(field, attr_name)
            if value is not None:
                state[attr_name.lower()] = str(value).strip()
        return state

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

    def _wait_for_validation_form(
        self,
        timeout_seconds: int = 30,
        raise_on_timeout: bool = True,
    ) -> bool:
        session = self.get_session()
        form_fields = (
            "wnd[0]/usr/ctxtP_FILE",
            "wnd[0]/usr/ctxtP_DATUM",
            "wnd[0]/usr/ctxtP_BUKRS",
            "wnd[0]/usr/txtP_WAERS",
            "wnd[0]/usr/txtP_FSEP",
            "wnd[0]/usr/chkP_TEST",
        )
        for _ in range(timeout_seconds):
            try:
                for field_id in form_fields:
                    session.findById(field_id)
                self.logger.info("SAP validation form detected.")
                return True
            except Exception:
                time.sleep(1)
        if raise_on_timeout:
            raise RuntimeError("SAP validation form was not ready after returning with F3.")
        return False

    def _is_results_grid_visible(self) -> bool:
        session = self.get_session()
        try:
            session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell")
            return True
        except Exception:
            return False

    def _handle_optional_popup(self) -> None:
        session = self.get_session()
        try:
            popup = session.findById("wnd[1]")
        except Exception:
            self.logger.info("No additional SAP popup detected during F3 fallback.")
            return

        try:
            popup_text = str(getattr(popup, "Text", "")).strip()
        except Exception:
            popup_text = "<unreadable>"

        self.logger.warning("SAP popup detected during F3 fallback | text=%s", popup_text)

        for button_id in ("tbar[0]/btn[0]", "tbar[0]/btn[12]"):
            try:
                popup.findById(button_id).press()
                self.logger.info("SAP popup button pressed during fallback | button_id=%s", button_id)
                return
            except Exception:
                continue

        self.logger.warning("SAP popup detected but no known fallback button could be pressed.")

    def _log_current_sap_state(self) -> None:
        session = self.get_session()
        state: dict[str, str] = {}

        try:
            state["wnd0_text"] = str(session.findById("wnd[0]").Text).strip()
        except Exception:
            state["wnd0_text"] = "<unreadable>"

        try:
            state["status_bar"] = str(session.findById("wnd[0]/sbar").Text).strip()
        except Exception:
            state["status_bar"] = "<unreadable>"

        try:
            state["active_window_name"] = str(getattr(session.ActiveWindow, "Name", "")).strip()
        except Exception:
            state["active_window_name"] = "<unreadable>"

        try:
            state["active_window_text"] = str(getattr(session.ActiveWindow, "Text", "")).strip()
        except Exception:
            state["active_window_text"] = "<unreadable>"

        self.logger.info("SAP current state during F3 fallback | %s", state)

    def _wait_for_login_screen(self, timeout_seconds: int = 30) -> None:
        session = self.get_session()
        login_fields = (
            "wnd[0]/usr/txtRSYST-BNAME",
            "wnd[0]/usr/pwdRSYST-BCODE",
            "wnd[0]/usr/txtRSYST-MANDT",
            "wnd[0]/usr/txtRSYST-LANGU",
        )
        for _ in range(timeout_seconds):
            try:
                for field_id in login_fields:
                    session.findById(field_id)
                return
            except Exception:
                time.sleep(1)
        raise RuntimeError("SAP login screen was not ready in time.")

    def _fill_login_form(self, username: str, password: str, client_number: str) -> None:
        session = self.get_session()
        self._wait_for_login_screen()

        username_field = session.findById("wnd[0]/usr/txtRSYST-BNAME")
        password_field = session.findById("wnd[0]/usr/pwdRSYST-BCODE")
        client_field = session.findById("wnd[0]/usr/txtRSYST-MANDT")
        language_field = session.findById("wnd[0]/usr/txtRSYST-LANGU")

        username_field.text = username
        password_field.text = password
        client_field.text = client_number
        language_field.text = self.settings.sap_language

        try:
            password_field.setFocus()
        except Exception:
            pass

        time.sleep(1)
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
        self.logger.info("SAP login submitted | client=%s", client_number)

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
    client.login("FMS")
    print("SAP session ready.")


def test_sap_tcode(system_group: str, tcode: str) -> None:
    """Standalone helper to verify TCode navigation after login."""
    client = SAPGuiClient()
    client.start()
    client.login(system_group)
    client.open_tcode(tcode)
    print(f"SAP TCode opened: {system_group} | {tcode}")


def test_sap_validation(csv_path: str) -> None:
    """Standalone helper to fill and execute the SAP validation screen for a CSV."""
    client = SAPGuiClient()
    payload = client.prepare_and_execute_validation(csv_path)
    print(f"SAP validation launched: {payload}")


def test_sap_validation_cycle(csv_path: str) -> None:
    """Standalone helper to validate a CSV in SAP until it is clean or blocked."""
    client = SAPGuiClient()
    result = client.validate_csv_until_clean(csv_path)
    print(f"SAP validation cycle result: {result}")


if __name__ == "__main__":
    test_sap_validation_cycle(
        r"runtime\outputs\Test\AP15_FMS_900010_USD_process_all_tickets_20260420_095349.csv"
    )