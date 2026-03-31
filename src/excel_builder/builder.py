from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path
from typing import Any

from src.processing.logic import normalize_numeric_code


AP15_HEADERS = [
    "Header Company code",
    "Vendor Code",
    "Invoice Number",
    "Invoice Date",
    "Source",
    "Distribution Type (DR/CR)",
    "Amount",
    "Currency USD/CAD",
    "G/L account Item Description",
    "Tax Type",
    "Company Code",
    "Profit Center 10 DIGITS",
    "Cost Center 10 DIGITS",
    "WBS",
    "Order",
    "Account",
    "Immediate Payment",
    "Special Handling Inst",
    "Paper Approval",
    "One Time vendor Name",
    "One Time vendor Street",
    "PO Box",
    "City",
    "State",
    "Zip",
    "Country",
    "Product Line",
    "Document type",
    "Tax code",
]


class AP15Builder:
    """Genera archivos CSV AP15 agrupados por VendorNum y Currency."""

    def __init__(self, output_dir: str) -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def build(
        self,
        records: list[dict[str, Any]],
        file_suffix: str = "",
    ) -> list[str]:
        grouped_records: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        for record in records:
            vendor_num = self._clean(record.get("VendorNum"))
            currency = self._clean(record.get("Currency"))
            grouped_records[(vendor_num, currency)].append(record)

        output_paths: list[str] = []
        suffix = f"_{file_suffix}" if file_suffix else ""
        for (vendor_num, currency), grouped in grouped_records.items():
            file_path = self.output_dir / f"AP15_{vendor_num}_{currency}{suffix}.csv"
            with file_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
                writer = csv.DictWriter(csv_file, fieldnames=AP15_HEADERS)
                writer.writeheader()
                for record in grouped:
                    writer.writerow(self._build_row(record))
            output_paths.append(str(file_path))
        return output_paths

    def _build_row(self, record: dict[str, Any]) -> dict[str, Any]:
        company_code = self._clean(record.get("CompanyCode"))
        vendor_num = self._clean(record.get("VendorNum"))
        invoice_num = self._clean(record.get("InvoiceNum"))
        invoice_date = self._clean(record.get("InvoiceDate"))
        currency = self._clean(record.get("Currency"))
        amount = record.get("Amount", "")
        cost_center = self._normalize_cost_center(record.get("CostCenter"))
        gl_account = self._normalize_account(record.get("GLAccount"))

        return {
            "Header Company code": company_code,
            "Vendor Code": vendor_num,
            "Invoice Number": invoice_num,
            "Invoice Date": invoice_date,
            "Source": "ITEM",
            "Distribution Type (DR/CR)": "DR",
            "Amount": amount,
            "Currency USD/CAD": currency,
            "G/L account Item Description": "",
            "Tax Type": "",
            "Company Code": company_code,
            "Profit Center 10 DIGITS": "",
            "Cost Center 10 DIGITS": cost_center,
            "WBS": self._clean(record.get("WBS")),
            "Order": "",
            "Account": gl_account,
            "Immediate Payment": "",
            "Special Handling Inst": "",
            "Paper Approval": "",
            "One Time vendor Name": self._clean(record.get("PayableTo")),
            "One Time vendor Street": self._clean(record.get("Address")),
            "PO Box": "",
            "City": self._clean(record.get("City")),
            "State": self._clean(record.get("State")),
            "Zip": self._clean(record.get("Zip")),
            "Country": self._clean(record.get("Country")),
            "Product Line": "",
            "Document type": "",
            "Tax code": "",
        }

    def _normalize_cost_center(self, value: Any) -> str:
        cleaned = self._clean(value)
        if cleaned in {"", "Attached"}:
            return ""
        normalized = normalize_numeric_code(cleaned, width=10)
        return "" if normalized == "Empty" else normalized

    def _normalize_account(self, value: Any) -> str:
        cleaned = self._clean(value).replace(" ", "")
        return "" if cleaned == "Empty" else cleaned

    def _clean(self, value: Any) -> str:
        if value is None:
            return ""
        cleaned = str(value).strip()
        return "" if cleaned == "Empty" else cleaned
