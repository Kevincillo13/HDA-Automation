import traceback
from datetime import datetime
from pathlib import Path
from typing import Any

from src.common.config import get_settings
from src.common.logger import get_logger
from src.common.models import TicketRecord
from src.common.run_context import get_run_id, start_run
from src.excel_builder.builder import AP15Builder
from src.hda_web.client import HDAClient
from src.hda_web.ticket_parser import extract_ticket_data
from src.processing.logic import apply_business_rules, validate_ticket_data


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
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    summary_timestamp = started_at.strftime("%Y%m%d_%H%M%S")
    summary_path = output_path / f"log_summary_{summary_timestamp}.txt"

    lines: list[str] = []
    lines.append("Daily Process Log Summary")
    lines.append(f"Run ID: {run_id}")
    lines.append(f"Started: {started_at.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Finished: {ended_at.strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Duration: {ended_at - started_at}")
    lines.append("")
    lines.append(
        f"The automation started at {started_at.strftime('%H:%M:%S')} and found "
        f"{len(one_time_checks)} 'OneTime Check' tickets."
    )
    lines.append("")

    lines.append("OneTime Check tickets found:")
    if one_time_checks:
        for index, ticket in enumerate(one_time_checks, start=1):
            lines.append(
                f"{index}. {ticket.ticket_id} | company={ticket.company} | "
                f"payment_method={ticket.payment_method} | subject={ticket.subject}"
            )
    else:
        lines.append("None")
    lines.append("")

    lines.append("Valid tickets:")
    if valid_results:
        for result in valid_results:
            ticket = result["ticket"]
            raw_data = result["raw_data"]
            processed_data = result["processed_data"]
            lines.append(
                f"- {ticket.ticket_id} | subject={ticket.subject} | company={ticket.company}"
            )
            lines.append(
                "  Raw: "
                f"Amount={raw_data.get('Amount')} | Invoice Number={raw_data.get('Invoice Number')} | "
                f"Invoice Date={raw_data.get('Invoice Date')} | Currency={raw_data.get('Currency')} | "
                f"Cost/Profit center={raw_data.get('Cost/Profit center')} | GL Account={raw_data.get('GL Account')}"
            )
            lines.append(
                "  Processed: "
                f"CompanyCode={processed_data.get('CompanyCode')} | VendorNum={processed_data.get('VendorNum')} | "
                f"InvoiceNum={processed_data.get('InvoiceNum')} | InvoiceDate={processed_data.get('InvoiceDate')} | "
                f"Amount={processed_data.get('Amount')} | Currency={processed_data.get('Currency')} | "
                f"CostCenter={processed_data.get('CostCenter')} | GLAccount={processed_data.get('GLAccount')} | "
                f"City={processed_data.get('City')} | State={processed_data.get('State')} | "
                f"Zip={processed_data.get('Zip')} | Country={processed_data.get('Country')}"
            )
            lines.append("")
    else:
        lines.append("None")
        lines.append("")

    lines.append("Invalid tickets:")
    if invalid_results:
        for result in invalid_results:
            ticket = result["ticket"]
            raw_data = result.get("raw_data", {})
            processed_data = result.get("processed_data", {})
            errors = result.get("errors", [])
            lines.append(
                f"- {ticket.ticket_id} | subject={ticket.subject} | company={ticket.company}"
            )
            if raw_data:
                lines.append(
                    "  Raw: "
                    f"Amount={raw_data.get('Amount')} | Invoice Number={raw_data.get('Invoice Number')} | "
                    f"Invoice Date={raw_data.get('Invoice Date')} | Currency={raw_data.get('Currency')} | "
                    f"Cost/Profit center={raw_data.get('Cost/Profit center')} | GL Account={raw_data.get('GL Account')}"
                )
            if processed_data:
                lines.append(
                    "  Processed: "
                    f"CompanyCode={processed_data.get('CompanyCode')} | VendorNum={processed_data.get('VendorNum')} | "
                    f"InvoiceNum={processed_data.get('InvoiceNum')} | InvoiceDate={processed_data.get('InvoiceDate')} | "
                    f"Amount={processed_data.get('Amount')} | Currency={processed_data.get('Currency')} | "
                    f"CostCenter={processed_data.get('CostCenter')} | GLAccount={processed_data.get('GLAccount')} | "
                    f"City={processed_data.get('City')} | State={processed_data.get('State')} | "
                    f"Zip={processed_data.get('Zip')} | Country={processed_data.get('Country')}"
                )
            lines.append("  Errors:")
            for error in errors:
                lines.append(f"  - {error}")
            lines.append("")
    else:
        lines.append("None")
        lines.append("")

    lines.append("Generated CSV files:")
    if generated_csvs:
        for csv_path in generated_csvs:
            lines.append(f"- {csv_path}")
    else:
        lines.append("None")
    lines.append("")

    summary_path.write_text("\n".join(lines), encoding="utf-8")
    return str(summary_path)


def _get_run_output_dir(base_output_dir: str, started_at: datetime) -> Path:
    return Path(base_output_dir) / started_at.strftime("%Y%m%d")


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
    run_id = get_run_id()
    started_at = datetime.now()
    run_output_dir = _get_run_output_dir(settings.output_dir, started_at)
    builder = AP15Builder(str(run_output_dir))

    logger.info("START PROCESS | run_id=%s", run_id)
    logger.info("Run outputs will be written to: %s", run_output_dir)
    valid_tickets: list[dict[str, Any]] = []
    invalid_tickets: list[dict[str, Any]] = []
    one_time_checks: list[TicketRecord] = []

    try:
        # --- LOGIN Y NAVEGACION ---
        client.start()
        client.log_debug_state("after_start")
        client.login()
        client.log_debug_state("after_login")
        client.click_payments_tile()
        client.log_debug_state("after_payments_click")

        # --- LECTURA Y FILTRADO DE TICKETS ---
        tickets = client.read_payment_grid_rows()
        logger.info("Grid extraction complete. Total tickets visible=%s", len(tickets))
        one_time_checks = [
            ticket for ticket in tickets if ticket.payment_method.strip() == "OneTime Check"
        ]
        logger.info(
            "Processing %s 'OneTime Check' tickets.", len(one_time_checks)
        )

        if not one_time_checks:
            logger.warning("No 'OneTime Check' tickets found. Exiting process.")
            return

        # --- BUCLE DE PROCESAMIENTO ---
        for i, ticket_to_process in enumerate(one_time_checks, start=1):
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
                    # client.reject_ticket(ticket_to_process.ticket_id, ", ".join(errors)) # Descomentar cuando la logica de rechazo este lista
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
        logger.info("--- PROCESSING COMPLETE ---")
        logger.info("Total tickets processed: %s", len(one_time_checks))
        logger.info("Valid tickets: %s", len(valid_tickets))
        logger.info("Invalid tickets: %s", len(invalid_tickets))

        generated_csvs: list[str] = []
        if valid_tickets:
            logger.info("--- VALID TICKETS COLLECTED ---")
            for valid_ticket in valid_tickets:
                logger.info(valid_ticket["processed_data"])
            generated_csvs = builder.build(
                [valid_ticket["processed_data"] for valid_ticket in valid_tickets],
                file_suffix=run_id,
            )
            logger.info("CSV files generated: %s", len(generated_csvs))
            for csv_path in generated_csvs:
                logger.info("Generated CSV: %s", csv_path)
        else:
            logger.info("No valid tickets collected, so no CSV files were generated.")

        summary_path = _write_human_summary(
            str(run_output_dir),
            run_id,
            started_at,
            datetime.now(),
            one_time_checks,
            valid_tickets,
            invalid_tickets,
            generated_csvs,
        )
        logger.info("Human-readable summary generated: %s", summary_path)

        # TODO: Implementar envio de email usando los CSV generados.

        logger.info("END PROCESS | status=success | run_id=%s", run_id)

    except KeyboardInterrupt:
        logger.warning("PROCESS INTERRUPTED BY USER | run_id=%s", run_id)
        raise
    except Exception as exc:
        logger.exception("An unhandled error stopped the process: %s", exc)
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
