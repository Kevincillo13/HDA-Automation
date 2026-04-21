from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from src.common.config import get_settings
from src.processing.logic import classify_mail_group, normalize_numeric_code

try:
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - depends on Windows environment
    win32com = None  # type: ignore[assignment]


class AP15Builder:
    """Genera AP15 a partir de template.xlsx y exporta CSVs desde ese layout."""

    def __init__(self, output_dir: str, template_path: str | None = None) -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        settings = get_settings()
        self.template_path = Path(template_path or settings.template_path)

    def build(
        self,
        records: list[dict[str, Any]],
        file_suffix: str = "",
    ) -> list[str]:
        grouped_records: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
        for record in records:
            mail_group = classify_mail_group(self._clean(record.get("CompanyCode")))
            vendor_num = self._clean(record.get("VendorNum"))
            currency = self._clean(record.get("Currency"))
            grouped_records[(mail_group, vendor_num, currency)].append(record)

        if not self.template_path.exists():
            raise FileNotFoundError(
                f"AP15 template not found: {self.template_path.resolve()}"
            )

        output_paths: list[str] = []
        suffix = f"_{file_suffix}" if file_suffix else ""
        for (mail_group, vendor_num, currency), grouped in grouped_records.items():
            base_name = f"AP15_{mail_group}_{vendor_num}_{currency}{suffix}"
            workbook_path = self.output_dir / f"{base_name}.xlsx"
            csv_path = self.output_dir / f"{base_name}.csv"

            workbook = load_workbook(self.template_path)
            worksheet = workbook.active
            self._fill_template_rows(worksheet, grouped)
            workbook.save(workbook_path)
            workbook.close()

            self._export_workbook_to_csv(workbook_path, csv_path)
            output_paths.append(str(csv_path))

        return output_paths

    def _fill_template_rows(self, worksheet: Any, records: list[dict[str, Any]]) -> None:
        for index, record in enumerate(records, start=2):
            company_code = self._clean(record.get("CompanyCode"))
            vendor_num = self._clean(record.get("VendorNum"))
            invoice_num = self._clean(record.get("InvoiceNum"))
            invoice_date = self._format_invoice_date(record.get("InvoiceDate"))
            amount = self._format_amount(record.get("Amount", ""))
            currency = self._clean(record.get("Currency"))
            center_value = self._normalize_cost_center(record.get("CostCenter"))
            gl_account = self._normalize_account(record.get("GLAccount"))
            profit_center, cost_center = self._route_center_fields(center_value, gl_account)

            worksheet[f"A{index}"] = company_code
            worksheet[f"B{index}"] = vendor_num
            worksheet[f"C{index}"] = invoice_num
            worksheet[f"D{index}"] = invoice_date
            worksheet[f"E{index}"] = "ITEM"
            worksheet[f"F{index}"] = "DR"
            if amount != "":
                worksheet[f"G{index}"] = float(amount)
                worksheet[f"G{index}"].number_format = "#,##0.00"
            else:
                worksheet[f"G{index}"] = ""
            worksheet[f"H{index}"] = currency
            worksheet[f"K{index}"] = company_code
            worksheet[f"L{index}"] = profit_center
            worksheet[f"M{index}"] = cost_center
            worksheet[f"N{index}"] = self._clean(record.get("WBS"))
            worksheet[f"O{index}"] = ""
            worksheet[f"P{index}"] = gl_account
            worksheet[f"Q{index}"] = ""
            worksheet[f"R{index}"] = ""
            worksheet[f"S{index}"] = ""
            worksheet[f"T{index}"] = self._clean(record.get("PayableTo"))
            worksheet[f"U{index}"] = self._clean(record.get("Address"))
            worksheet[f"V{index}"] = ""
            worksheet[f"W{index}"] = self._clean(record.get("City"))
            worksheet[f"X{index}"] = self._clean(record.get("State"))
            worksheet[f"Y{index}"] = self._clean(record.get("Zip"))
            worksheet[f"Z{index}"] = self._clean(record.get("Country"))
            worksheet[f"AA{index}"] = ""
            worksheet[f"AB{index}"] = ""
            worksheet[f"AC{index}"] = ""

    def _export_workbook_to_csv(self, workbook_path: Path, csv_path: Path) -> None:
        if win32com is not None:
            self._export_workbook_to_csv_via_excel(workbook_path, csv_path)
            return

        # Fallback only when Excel COM is unavailable.
        self._export_workbook_to_csv_via_python(workbook_path, csv_path)

    def _export_workbook_to_csv_via_excel(self, workbook_path: Path, csv_path: Path) -> None:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = None
        try:
            workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
            target_path = str(csv_path.resolve())
            if csv_path.exists():
                csv_path.unlink()
            # 62 = xlCSVUTF8. This matches the manual "CSV UTF-8" save flow well.
            workbook.SaveAs(target_path, FileFormat=62, Local=True)
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()

    def _export_workbook_to_csv_via_python(self, workbook_path: Path, csv_path: Path) -> None:
        workbook = load_workbook(workbook_path, data_only=True)
        worksheet = workbook.active
        max_row = self._find_last_non_empty_row(worksheet)
        max_column = 29

        with csv_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
            writer = csv.writer(csv_file, lineterminator="\r\n")
            for row_index in range(1, max_row + 1):
                row_values: list[str] = []
                for column_index in range(1, max_column + 1):
                    value = worksheet.cell(row=row_index, column=column_index).value
                    row_values.append(self._serialize_cell_value(value))
                writer.writerow(row_values)

        workbook.close()

    def _find_last_non_empty_row(self, worksheet: Any) -> int:
        last_row = 1
        for row_index in range(1, worksheet.max_row + 1):
            if any(
                worksheet.cell(row=row_index, column=column_index).value not in (None, "")
                for column_index in range(1, 30)
            ):
                last_row = row_index
        return last_row

    def _serialize_cell_value(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, float):
            return f"{value:.2f}"
        return str(value)

    def _normalize_cost_center(self, value: Any) -> str:
        cleaned = self._clean(value)
        if cleaned in {"", "Attached"}:
            return ""
        normalized = normalize_numeric_code(cleaned, width=10)
        return "" if normalized == "Empty" else normalized

    def _normalize_account(self, value: Any) -> str:
        cleaned = self._clean(value).replace(" ", "")
        return "" if cleaned == "Empty" else cleaned

    def _route_center_fields(self, center_value: str, account_value: str) -> tuple[str, str]:
        normalized_account = self._clean(account_value).upper()
        if not center_value:
            return "", ""

        profit_prefixes = ("11", "12", "13", "P1", "P2", "P3")
        cost_prefixes = ("14", "15", "16", "P4", "P5", "P6")

        if normalized_account.startswith(profit_prefixes):
            return center_value, ""
        if normalized_account.startswith(cost_prefixes):
            return "", center_value

        return "", center_value

    def _format_invoice_date(self, value: Any) -> str:
        cleaned = self._clean(value)
        if not cleaned:
            return ""

        parts = cleaned.split("/")
        if len(parts) != 3:
            return cleaned

        try:
            month = str(int(parts[0]))
            day = str(int(parts[1]))
            year = str(int(parts[2]))
            return f"{month}/{day}/{year}"
        except ValueError:
            return cleaned

    def _format_amount(self, value: Any) -> str:
        if value in (None, ""):
            return ""
        try:
            return f"{float(value):.2f}"
        except (TypeError, ValueError):
            return self._clean(value)

    def _clean(self, value: Any) -> str:
        if value is None:
            return ""
        cleaned = str(value).strip()
        return "" if cleaned == "Empty" else cleaned
