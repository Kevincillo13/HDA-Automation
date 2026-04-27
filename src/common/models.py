from dataclasses import dataclass, field
from typing import List


@dataclass
class TicketRecord:
    ticket_id: str
    created: str = ""
    payment_method: str = ""
    subject: str = ""
    company: str = ""
    ticket_type: str = ""
    hda_status: str = ""
    pdf_path: str = ""
    excel_path: str = ""
    validation_errors: List[str] = field(default_factory=list)
    status: str = "pending"


@dataclass
class RunSummary:
    processed: int = 0
    rejected: int = 0
    generated_excels: int = 0
    emailed: int = 0
    failed: int = 0
