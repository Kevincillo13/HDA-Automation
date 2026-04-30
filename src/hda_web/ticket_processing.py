import traceback
import time
from collections import defaultdict
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Any

from src.common.config import get_settings
from src.common.logger import get_logger
from src.common.models import TicketRecord
from src.common.run_context import get_run_id, start_run
from src.common.system import kill_processes
from src.excel_builder.builder import AP15Builder
from src.hda_web.client import HDAClient
from src.hda_web.ticket_parser import extract_ticket_data
from src.mailer.client import SMTPMailClient
from src.processing.logic import apply_business_rules, validate_ticket_data
from src.sap.client import SAPGuiClient


def _write_human_summary(
    output_dir: str,
    run_id: str,
    started_at: datetime,
    ended_at: datetime,
    one_time_checks: list[TicketRecord],
    valid_results: list[dict[str, Any]],
    invalid_results: list[dict[str, Any]],
    generated_csvs: list[str],
) -> str:
    """Writes a human-readable summary of the process results after all stages (HDA + SAP)."""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    summary_timestamp = started_at.strftime("%Y%m%d_%H%M%S")
    summary_path = output_path / f"log_summary_{summary_timestamp}.txt"

    lines: list[str] = []
    lines.append("Daily Process Log Summary")
    lines.append("=" * 30)
    lines.append(f"Run ID: {run_id}")
    lines.append(f"Started: {started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Finished: {ended_at.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Duration: {ended_at - started_at}")
    lines.append("")
    lines.append(f"1. Total OneTime Check detected in HDA Grid: {len(one_time_checks)}")
    lines.append(f"2. Passed HDA Local Validation: {len(valid_results) + len([r for r in invalid_results if 'Suspended by SAP' in str(r.get('errors', []))])}")
    lines.append(f"3. Passed SAP Final Validation: {len(valid_results)}")
    lines.append("")
    lines.append(f"FINAL RESULT: {len(valid_results)} tickets ready to pay, {len(invalid_results)} tickets suspended.")
    lines.append("")

    lines.append("--- ONE-TIME CHECK TICKETS DETECTED ---")
    if one_time_checks:
        for index, ticket in enumerate(one_time_checks, start=1):
            lines.append(
                f"{index}. {ticket.ticket_id} | company={ticket.company} | "
                f"payment_method={ticket.payment_method} | subject={ticket.subject}"
            )
    else:
        lines.append("None")
    lines.append("")

    lines.append("--- FINAL VALID TICKETS (Passed SAP Validation) ---")
    if valid_results:
        for result in valid_results:
            ticket = result["ticket"]
            proc = result["processed_data"]
            lines.append(f"- {ticket.ticket_id} | subject={ticket.subject} | company={ticket.company}")
            lines.append(f"  Amount: {proc.get('Amount')} {proc.get('Currency')} | Invoice: {proc.get('InvoiceNum')}")
            lines.append("")
    else:
        lines.append("None")
        lines.append("")

    lines.append("--- INVALID OR SUSPENDED TICKETS (Local or SAP) ---")
    if invalid_results:
        for result in invalid_results:
            ticket = result["ticket"]
            errors = result.get("errors", [])
            lines.append(f"- {ticket.ticket_id} | subject={ticket.subject} | company={ticket.company}")
            lines.append("  Reasons for suspension:")
            for error in errors:
                # Si el error tiene saltos de línea (como los de SAP), indentamos cada línea
                error_lines = str(error).splitlines()
                if error_lines:
                    lines.append(f"    * {error_lines[0]}")
                    for subline in error_lines[1:]:
                        lines.append(f"      {subline}")
            lines.append("")
    else:
        lines.append("None")
        lines.append("")

    lines.append("--- GENERATED FINAL CSV FILES ---")
    if generated_csvs:
        for csv_path in generated_csvs:
            lines.append(f"- {Path(csv_path).name}")
    else:
        lines.append("None")

    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return str(summary_path)
    lines.append("")

    summary_path.write_text("\n".join(lines), encoding="utf-8")
    return str(summary_path)


def _get_run_output_dir(base_output_dir: str, started_at: datetime) -> Path:
    return Path(base_output_dir) / started_at.strftime("%Y%m%d")


def _build_typed_email_subject(
    settings: Any, started_at: datetime, mail_type: str
) -> str:
    prefix = settings.mail_subject_prefix.strip()
    base_subject = f"HDA {mail_type} - {started_at.strftime('%Y-%m-%d')}"
    return f"{prefix} {base_subject}".strip() if prefix else base_subject


def _build_group_email_body(
    run_id: str,
    started_at: datetime,
    mail_group: str,
    attachments: list[str],
) -> str:
    file_names = [Path(attachment).name for attachment in attachments]
    lines = [
        "Hello,",
        "",
        f"Attached are the HDA AP15 {mail_group} file(s) generated by the automation.",
        "",
        f"Run ID: {run_id}",
        f"Started: {started_at.strftime('%Y-%m-%d %H:%M:%S')}",
        f"{mail_group} CSV files attached: {len(attachments)}",
        "",
        "Attached files:",
    ]
    for file_name in file_names:
        lines.append(f"- {file_name}")
    return "\n".join(lines)


def _build_summary_email_body(summary_text: str, run_id: str, started_at: datetime) -> str:
    return "\n".join(
        [
            "Hello,",
            "",
            "Below is the daily HDA automation summary.",
            "",
            f"Run ID: {run_id}",
            f"Started: {started_at.strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            summary_text,
        ]
    )


def _build_summary_email_html(
    run_id: str,
    started_at: datetime,
    ended_at: datetime,
    one_time_checks: list[TicketRecord],
    valid_results: list[dict[str, Any]],
    invalid_results: list[dict[str, Any]],
    generated_csvs: list[str],
) -> str:
    duration = ended_at - started_at

    def render_ticket_rows(results: list[dict[str, Any]], include_errors: bool) -> str:
        if not results:
            label = "No tickets in this section."
            return (
                "<tr><td colspan='4' style='padding:12px;border:1px solid #d9dee8;"
                "color:#5b6472;'>"
                f"{escape(label)}</td></tr>"
            )

        rows: list[str] = []
        for result in results:
            ticket = result["ticket"]
            processed = result.get("processed_data", {})
            errors = result.get("errors", [])
            detail_lines = [
                f"Company: {ticket.company or 'N/A'}",
                f"Subject: {ticket.subject or 'N/A'}",
                f"Invoice: {processed.get('InvoiceNum', 'N/A')}",
                f"Amount: {processed.get('Amount', 'N/A')} {processed.get('Currency', '')}".strip(),
                f"Vendor: {processed.get('VendorNum', 'N/A')}",
                f"GL: {processed.get('GLAccount', 'N/A')}",
            ]
            if include_errors and errors:
                # Unir múltiples errores con saltos de línea y convertir \n a <br>
                err_text = "<br>".join(escape(str(error)).replace("\n", "<br>") for error in errors)
                detail_lines.append(f"Errors: {err_text}")
            details = "<br>".join(detail_lines) # Quité el escape global porque err_text ya está escapado
            rows.append(
                "<tr>"
                f"<td style='padding:12px;border:1px solid #d9dee8;vertical-align:top;'>{escape(ticket.ticket_id)}</td>"
                f"<td style='padding:12px;border:1px solid #d9dee8;vertical-align:top;'>{escape(ticket.company or 'N/A')}</td>"
                f"<td style='padding:12px;border:1px solid #d9dee8;vertical-align:top;'>{escape(ticket.subject or 'N/A')}</td>"
                f"<td style='padding:12px;border:1px solid #d9dee8;vertical-align:top;'>{details}</td>"
                "</tr>"
            )
        return "".join(rows)

    def render_simple_list(items: list[str], empty_text: str) -> str:
        if not items:
            return f"<p style='margin:0;color:#5b6472;'>{escape(empty_text)}</p>"
        return "<ul style='margin:0;padding-left:18px;'>" + "".join(
            f"<li style='margin:4px 0;'>{escape(item)}</li>" for item in items
        ) + "</ul>"

    metrics = [
        ("Run ID", run_id),
        ("Started", started_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Finished", ended_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Duration", str(duration)),
        ("OneTime Check tickets", str(len(one_time_checks))),
        ("Valid tickets", str(len(valid_results))),
        ("Invalid tickets", str(len(invalid_results))),
        ("CSV files generated", str(len(generated_csvs))),
    ]

    metric_cards = "".join(
        "<div style='display:inline-block;width:220px;margin:0 12px 12px 0;padding:14px 16px;"
        "background:#f5f7fb;border:1px solid #d9dee8;border-radius:10px;vertical-align:top;'>"
        f"<div style='font-size:12px;color:#5b6472;text-transform:uppercase;letter-spacing:.04em;'>{escape(label)}</div>"
        f"<div style='margin-top:6px;font-size:18px;font-weight:700;color:#122033;'>{escape(value)}</div>"
        "</div>"
        for label, value in metrics
    )

    one_time_items = [
        f"{ticket.ticket_id} | {ticket.company} | {ticket.subject}" for ticket in one_time_checks
    ]
    csv_items = [Path(csv_path).name for csv_path in generated_csvs]

    return f"""\
<html>
  <body style="margin:0;padding:24px;background:#eef3f8;font-family:Segoe UI,Arial,sans-serif;color:#122033;">
    <div style="max-width:960px;margin:0 auto;background:#ffffff;border:1px solid #d9dee8;border-radius:16px;overflow:hidden;">
      <div style="padding:24px 28px;background:#183153;color:#ffffff;">
        <div style="font-size:13px;letter-spacing:.08em;text-transform:uppercase;opacity:.85;">HDA Automation</div>
        <h1 style="margin:10px 0 0;font-size:28px;line-height:1.2;">Daily Summary</h1>
        <p style="margin:10px 0 0;font-size:15px;opacity:.92;">This email summarizes the latest OneTime Check automation run.</p>
      </div>
      <div style="padding:28px;">
        <h2 style="margin:0 0 16px;font-size:18px;color:#122033;">Overview</h2>
        <div>{metric_cards}</div>

        <h2 style="margin:24px 0 12px;font-size:18px;color:#122033;">OneTime Check Tickets Found</h2>
        {render_simple_list(one_time_items, "No OneTime Check tickets were found.")}

        <h2 style="margin:24px 0 12px;font-size:18px;color:#122033;">Valid Tickets</h2>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead>
            <tr style="background:#f5f7fb;">
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Ticket</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Company</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Subject</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Details</th>
            </tr>
          </thead>
          <tbody>
            {render_ticket_rows(valid_results, include_errors=False)}
          </tbody>
        </table>

        <h2 style="margin:24px 0 12px;font-size:18px;color:#122033;">Invalid Tickets</h2>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead>
            <tr style="background:#fff4f2;">
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Ticket</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Company</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Subject</th>
              <th style="padding:12px;border:1px solid #d9dee8;text-align:left;">Details</th>
            </tr>
          </thead>
          <tbody>
            {render_ticket_rows(invalid_results, include_errors=True)}
          </tbody>
        </table>

        <h2 style="margin:24px 0 12px;font-size:18px;color:#122033;">Generated CSV Files</h2>
        {render_simple_list(csv_items, "No CSV files were generated in this run.")}
      </div>
    </div>
  </body>
</html>
"""


def _build_error_email_body(
    run_id: str,
    started_at: datetime,
    current_stage: str,
    exc: Exception,
    log_path: str,
) -> str:
    return "\n".join(
        [
            "Hello,",
            "",
            "The HDA automation failed.",
            "",
            f"Run ID: {run_id}",
            f"Started: {started_at.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Failed during: {current_stage}",
            f"Exception type: {type(exc).__name__}",
            f"Error message: {exc}",
            f"Technical log: {log_path}",
            "",
            "Please review the attached technical log for more details.",
            "",
            "Traceback:",
            traceback.format_exc(),
        ]
    )


def _get_log_path(settings: Any, run_id: str) -> str:
    return str(Path(settings.log_dir) / f"{run_id}.log")


def _get_mail_group_from_csv_path(csv_path: str) -> str | None:
    csv_name = Path(csv_path).name.upper()
    if "_FMS_" in csv_name:
        return "FMS"
    if "_AFS_" in csv_name:
        return "AFS"
    return None


def _group_csvs_by_mail_group(generated_csvs: list[str]) -> dict[str, list[str]]:
    grouped = {"FMS": [], "AFS": []}
    for csv_path in generated_csvs:
        mail_group = _get_mail_group_from_csv_path(csv_path)
        if mail_group:
            grouped[mail_group].append(csv_path)
    return grouped


def _resolve_mail_recipients(settings: Any, mail_type: str) -> list[str]:
    recipient_by_type = {
        "FMS": settings.mail_fms_recipient or settings.mail_usd_recipient,
        "AFS": settings.mail_afs_recipient or settings.mail_cad_recipient,
        "SUMMARY": settings.mail_summary_recipient,
        "ERROR": settings.mail_error_recipient,
    }
    preferred_recipient = recipient_by_type.get(mail_type, "")
    if settings.mail_test_recipient:
        preferred_recipient = settings.mail_test_recipient
    elif not preferred_recipient:
        preferred_recipient = ",".join(
            recipient
            for recipient in [
                settings.mail_primary_recipient,
                settings.mail_secondary_recipient,
            ]
            if recipient
        )
    normalized = preferred_recipient.replace(";", ",")
    return [recipient.strip() for recipient in normalized.split(",") if recipient.strip()]

def _collect_one_time_checks_with_retry(
    client: HDAClient,
    logger: Any,
    attempts: int = 3,
    wait_seconds: float = 4.0,
) -> tuple[list[TicketRecord], list[TicketRecord]]:
    last_tickets: list[TicketRecord] = []
    one_time_checks: list[TicketRecord] = []

    for attempt in range(1, attempts + 1):
        if attempt > 1:
            logger.warning(
                "Retrying Payments grid extraction because no 'OneTime Check' tickets were detected on attempt %s/%s.",
                attempt - 1,
                attempts,
            )
            try:
                client.click_payments_tile()
            except Exception as exc:
                logger.warning(
                    "Payments tile reactivation failed before retry %s/%s: %s",
                    attempt,
                    attempts,
                    exc,
                )
            client.log_debug_state(f"payments_retry_{attempt}")
            time.sleep(wait_seconds)

        tickets = client.read_payment_grid_rows()
        last_tickets = tickets
        
        # --- EL FILTRO CORREGIDO Y BLINDADO ---
        # Usamos "in" y eliminamos caracteres invisibles (\xa0) para evitar que ExtJS nos engañe.
        # Además, validamos explícitamente los estados permitidos.
        one_time_checks = [
            ticket for ticket in tickets 
            if "onetime check" in ticket.payment_method.replace('\xa0', ' ').strip().lower()
            and ticket.status.replace('\xa0', ' ').strip().lower() in ["open", "assigned", "suspended"]
        ]

        empty_payment_method = len(
            [ticket for ticket in tickets if not ticket.payment_method.strip()]
        )
        logger.info(
            "Payments extraction attempt %s/%s | total_tickets=%s | one_time_checks=%s | empty_payment_method=%s",
            attempt,
            attempts,
            len(tickets),
            len(one_time_checks),
            empty_payment_method,
        )

        if one_time_checks:
            return last_tickets, one_time_checks

        if tickets:
            preview = [
                f"{ticket.ticket_id}:{ticket.payment_method or '<empty>'}"
                for ticket in tickets[:10]
            ]
            logger.warning(
                "No 'OneTime Check' tickets detected on attempt %s/%s despite visible rows. Sample payment methods: %s",
                attempt,
                attempts,
                preview,
            )
        else:
            logger.warning(
                "No tickets were visible on Payments extraction attempt %s/%s.",
                attempt,
                attempts,
            )

    return last_tickets, one_time_checks

def process_all_tickets() -> None:
    """
    Orquesta el proceso completo de HDA:
    1. Inicia sesión.
    2. Lee todos los tickets "OneTime Check".
    3. Itera sobre cada ticket, lo procesa, valida y actúa.
    4. Acumula los resultados.
    """
    start_run("process_all_tickets")
    logger = get_logger("process_all_tickets")
    settings = get_settings()
    client = HDAClient()
    mail_client = SMTPMailClient(settings)
    run_id = get_run_id()
    started_at = datetime.now()
    run_output_dir = _get_run_output_dir(settings.output_dir, started_at)
    log_path = _get_log_path(settings, run_id)
    builder = AP15Builder(str(run_output_dir))
    current_stage = "initialization"

    # --- LIMPIEZA INICIAL ---
    logger.info("Cleaning up environment before starting...")
    kill_processes(["msedge.exe", "msedgedriver.exe", "saplogon.exe", "sapgui.exe"])
    time.sleep(2)  # Dar tiempo al OS de liberar los locks de los perfiles de Edge

    logger.info("START PROCESS | run_id=%s", run_id)
    logger.info("Run outputs will be written to: %s", run_output_dir)
    valid_tickets: list[dict[str, Any]] = []
    invalid_tickets: list[dict[str, Any]] = []
    one_time_checks: list[TicketRecord] = []

    try:
        # --- LOGIN Y NAVEGACION ---
        current_stage = "browser startup"
        client.start()
        client.log_debug_state("after_start")
        current_stage = "login"
        client.login()
        client.log_debug_state("after_login")
        current_stage = "navigation to Payments"
        client.click_payments_tile()
        client.log_debug_state("after_payments_click")

        # --- LECTURA Y FILTRADO DE TICKETS ---
        current_stage = "ticket grid extraction"
        tickets, one_time_checks = _collect_one_time_checks_with_retry(client, logger)
        logger.info("Grid extraction complete. Total tickets visible=%s", len(tickets))
        logger.info(
            "Processing %s 'OneTime Check' tickets.", len(one_time_checks)
        )

        if not one_time_checks:
            logger.warning(
                "No 'OneTime Check' tickets found after retrying Payments extraction. Exiting process."
            )
            return

        # --- BUCLE DE PROCESAMIENTO ---
        for i, ticket_to_process in enumerate(one_time_checks, start=1):
            current_stage = f"processing ticket {ticket_to_process.ticket_id}"
            logger.info(
                "--- Processing ticket %s/%s: %s ---",
                i,
                len(one_time_checks),
                ticket_to_process.ticket_id,
            )
            try:
                # 1. Abrir ticket
                client.open_ticket_by_id(ticket_to_process.ticket_id)
                client.take_screenshot(f"ticket_{ticket_to_process.ticket_id}_open")

                # 2. Extraer y procesar datos
                raw_data = extract_ticket_data(client.driver)
                raw_data["Id"] = ticket_to_process.ticket_id
                raw_data["Created"] = ticket_to_process.created or raw_data.get(
                    "Created",
                    "Empty",
                )
                logger.info("Raw data extracted: %s", raw_data)
                processed_data = apply_business_rules(raw_data)
                logger.info("Processed data: %s", processed_data)

                # 3. Validar
                errors = validate_ticket_data(processed_data)

                # 4. Tomar accion
                if errors:
                    logger.warning("Ticket %s is INVALID.", ticket_to_process.ticket_id)
                    for error in errors:
                        logger.warning("- %s", error)
                    invalid_tickets.append(
                        {
                            "ticket": ticket_to_process,
                            "raw_data": raw_data,
                            "processed_data": processed_data,
                            "errors": errors,
                        }
                    )
                    # client.suspend_ticket(ticket_to_process.ticket_id, ", ".join(errors)) # Descomentar cuando la logica de suspensión este lista
                else:
                    logger.info("Ticket %s is VALID.", ticket_to_process.ticket_id)
                    valid_tickets.append(
                        {
                            "ticket": ticket_to_process,
                            "raw_data": raw_data,
                            "processed_data": processed_data,
                        }
                    )

            except Exception as e:
                logger.error(
                    "Failed to process ticket %s: %s",
                    ticket_to_process.ticket_id,
                    e,
                    exc_info=True,
                )
                invalid_tickets.append(
                    {
                        "ticket": ticket_to_process,
                        "raw_data": {},
                        "processed_data": {},
                        "errors": ["Unhandled exception during processing"],
                    }
                )
                client.take_screenshot(f"ticket_{ticket_to_process.ticket_id}_error")
            finally:
                # 5. Cerrar ticket (siempre se ejecuta)
                client.close_active_ticket_tab()

        # --- FINALIZACION Y RESUMEN ---
        current_stage = "final summary generation"
        logger.info("--- PROCESSING COMPLETE ---")
        logger.info("Total tickets processed: %s", len(one_time_checks))
        logger.info("Valid tickets: %s", len(valid_tickets))
        logger.info("Invalid (Suspended) tickets: %s", len(invalid_tickets))

        generated_csvs: list[str] = []
        if valid_tickets:
            current_stage = "csv generation"
            logger.info("--- VALID TICKETS COLLECTED ---")
            for valid_ticket in valid_tickets:
                logger.info(valid_ticket["processed_data"])
            
            # 1. Generación inicial de CSVs (fragmentados por Vendor)
            candidate_csvs = builder.build(
                [valid_ticket["processed_data"] for valid_ticket in valid_tickets],
                file_suffix=run_id,
            )
            logger.info("Candidate CSV files generated: %s", len(candidate_csvs))

            # 2. Validación SAP
            current_stage = "SAP validation"
            final_valid_csvs_by_group = defaultdict(list)
            sap_suspended_invoices = set()
            sap_all_suspension_reasons = {}

            for csv_path in candidate_csvs:
                logger.info("--- Starting SAP Validation for: %s ---", Path(csv_path).name)
                
                # Identificar grupo y vendor desde el nombre del archivo original
                mail_group = _get_mail_group_from_csv_path(csv_path) or "GRP"
                parts = Path(csv_path).name.split("_")
                vendor = "UNK"
                currency = "UNK"
                if len(parts) >= 4:
                    vendor = parts[2]
                    currency = parts[3]

                sap_client = SAPGuiClient(settings)
                try:
                    # Sufijo corto para evitar errores de longitud en SAP
                    unique_suffix = f"{mail_group}_{vendor}"
                    sap_result = sap_client.validate_csv_until_clean(csv_path, retry_suffix=unique_suffix)
                    
                    # Recolectar suspensiones y sus motivos siempre
                    sap_suspended_invoices.update(sap_result.get("all_suspended_invoices", []))
                    sap_all_suspension_reasons.update(sap_result.get("suspension_reasons", {}))

                    if sap_result["status"] == "clean":
                        clean_path = Path(sap_result["final_csv_path"])
                        # VALIDACIÓN CRÍTICA: ¿Tiene datos el archivo limpio?
                        _, rows = sap_client._read_csv_rows(str(clean_path))
                        if rows:
                            final_valid_csvs_by_group[(mail_group, currency)].append(str(clean_path))
                            logger.info("Clean CSV with %s rows added to final pool.", len(rows))
                        else:
                            logger.warning("Clean CSV for vendor %s is empty, skipping.", vendor)
                except Exception as sap_exc:
                    logger.error("Technical error during SAP validation for %s: %s", vendor, sap_exc)
                finally:
                    sap_client.close()

            # 3. Mover tickets suspendidos por SAP a la lista de inválidos para el reporte
            if sap_suspended_invoices:
                logger.info("SAP suspended %s invoices. Updating counts.", len(sap_suspended_invoices))
                still_valid = []
                for entry in valid_tickets:
                    inv = str(entry["processed_data"].get("InvoiceNum", "")).strip()
                    if inv in sap_suspended_invoices:
                        reasons = sap_all_suspension_reasons.get(inv, [])
                        msg = "Suspended by SAP validation."
                        if reasons:
                            msg += "\n" + "\n".join(f"- {r}" for r in reasons)
                        entry["errors"] = entry.get("errors", []) + [msg]
                        invalid_tickets.append(entry)
                    else:
                        still_valid.append(entry)
                valid_tickets = still_valid

            # 4. Consolidación Final (Solo si hay datos)
            current_stage = "final consolidation"
            generated_csvs = []
            
            for (group, curr), paths in final_valid_csvs_by_group.items():
                if not paths:
                    continue
                
                # Nombre de archivo final según requerimiento: AP15_GROUP_FINAL_YYMMDD
                date_str = started_at.strftime("%y%m%d")
                consolidated_name = f"AP15_{group}_FINAL_{date_str}_{curr}"
                
                consolidated_path = builder.merge_csvs(paths, consolidated_name)
                generated_csvs.append(consolidated_path)
                logger.info("FINAL Consolidated CSV generated: %s", consolidated_path)

        else:
            logger.info("No valid tickets passed HDA validation.")
            
        # --- NUEVO PASO: SUSPENSIÓN AUTOMÁTICA EN EL PORTAL ---
        if invalid_tickets:
            current_stage = "automatic suspension in HDA"
            logger.info("--- STARTING AUTOMATIC SUSPENSION FOR %s TICKETS ---", len(invalid_tickets))
            for item in invalid_tickets:
                ticket_id = item["ticket"].ticket_id
                reasons = item.get("errors", ["Unknown validation failure"])
                
                try:
                    logger.info("Suspending ticket %s in HDA...", ticket_id)
                    client.open_ticket_by_id(ticket_id)
                    client.suspend_ticket_ui(ticket_id, reasons)
                    logger.info("Ticket %s successfully suspended.", ticket_id)
                except Exception as susp_exc:
                    logger.error("Failed to suspend ticket %s: %s", ticket_id, susp_exc)
                    client.take_screenshot(f"suspend_fail_{ticket_id}")
                finally:
                    client.close_active_ticket_tab()
            
            logger.info("--- AUTOMATIC SUSPENSION PROCESS COMPLETE ---")

        # 5. Generar reporte con terminología clara
        summary_path = _write_human_summary(
            str(run_output_dir),
            run_id,
            started_at,
            datetime.now(),
            one_time_checks, # Estos son todos los encontrados
            valid_tickets,    # Estos son los que pasaron SAP
            invalid_tickets,  # Estos incluyen suspensiones HDA y suspensiones SAP
            generated_csvs,
        )
        logger.info("Human-readable summary generated: %s", summary_path)

        if settings.smtp_host and (
            settings.mail_test_recipient
            or settings.mail_primary_recipient
            or settings.mail_secondary_recipient
            or settings.mail_fms_recipient
            or settings.mail_afs_recipient
            or settings.mail_usd_recipient
            or settings.mail_cad_recipient
            or settings.mail_summary_recipient
            or settings.mail_error_recipient
        ):
            current_stage = "email sending"
            mail_groups = _group_csvs_by_mail_group(generated_csvs)
            for mail_group, group_csvs in mail_groups.items():
                if not group_csvs:
                    continue
                try:
                    mail_result = mail_client.send_message(
                        subject=_build_typed_email_subject(settings, started_at, f"AP15 {mail_group}"),
                        body=_build_group_email_body(
                            run_id=run_id,
                            started_at=started_at,
                            mail_group=mail_group,
                            attachments=group_csvs,
                        ),
                        attachments=group_csvs,
                        recipients=_resolve_mail_recipients(settings, mail_group),
                    )
                    logger.info(
                        "%s CSV email sent | recipients=%s | bcc=%s | attachments=%s",
                        mail_group,
                        mail_result.recipients,
                        mail_result.bcc,
                        len(mail_result.attachments),
                    )
                except Exception as mail_exc:
                    logger.exception("%s CSV email failed to send: %s", mail_group, mail_exc)

            try:
                ended_at = datetime.now()
                summary_text = Path(summary_path).read_text(encoding="utf-8")
                summary_mail_result = mail_client.send_message(
                    subject=_build_typed_email_subject(settings, started_at, "Automation Summary"),
                    body=_build_summary_email_body(summary_text, run_id, started_at),
                    html_body=_build_summary_email_html(
                        run_id=run_id,
                        started_at=started_at,
                        ended_at=ended_at,
                        one_time_checks=one_time_checks,
                        valid_results=valid_tickets,
                        invalid_results=invalid_tickets,
                        generated_csvs=generated_csvs,
                    ),
                    attachments=[],
                    recipients=_resolve_mail_recipients(settings, "SUMMARY"),
                )
                logger.info(
                    "Summary email sent | recipients=%s | bcc=%s | attachments=%s",
                    summary_mail_result.recipients,
                    summary_mail_result.bcc,
                    len(summary_mail_result.attachments),
                )
            except Exception as mail_exc:
                logger.exception("Summary email failed to send: %s", mail_exc)
        else:
            logger.info("Email sending skipped because SMTP or recipients are not configured.")

        logger.info("END PROCESS | status=success | run_id=%s", run_id)

    except KeyboardInterrupt:
        logger.warning("PROCESS INTERRUPTED BY USER | run_id=%s", run_id)
        raise
    except Exception as exc:
        logger.exception("An unhandled error stopped the process: %s", exc)
        if settings.smtp_host and (
            settings.mail_test_recipient
            or settings.mail_primary_recipient
            or settings.mail_secondary_recipient
            or settings.mail_error_recipient
        ):
            try:
                error_mail_result = mail_client.send_message(
                    subject=_build_typed_email_subject(settings, started_at, "Automation FAILED"),
                    body=_build_error_email_body(
                        run_id=run_id,
                        started_at=started_at,
                        current_stage=current_stage,
                        exc=exc,
                        log_path=log_path,
                    ),
                    attachments=[log_path],
                    recipients=_resolve_mail_recipients(settings, "ERROR"),
                )
                logger.info(
                    "Failure email sent | recipients=%s | bcc=%s | attachments=%s",
                    error_mail_result.recipients,
                    error_mail_result.bcc,
                    len(error_mail_result.attachments),
                )
            except Exception as mail_exc:
                logger.exception("Failure email failed to send: %s", mail_exc)
        try:
            client.log_debug_state("exception")
            client.log_browser_console("exception")
            client.take_screenshot("99_exception")
            client.save_page_source("99_exception")
        except Exception:
            logger.error(
                "Failed to collect extra debug evidence.\n%s", traceback.format_exc()
            )
        logger.error("END PROCESS | status=failed | run_id=%s", run_id)
        raise
    finally:
        if settings.browser_keep_open and client.driver:
            logger.info(
                "Browser will be kept open. Press Enter to exit.",
            )
            try:
                input("Press Enter to close the browser...")
            except EOFError:
                pass
        client.close()


if __name__ == "__main__":
    process_all_tickets()
