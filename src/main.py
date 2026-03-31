from src.hda_web.ticket_processing import process_all_tickets


def main() -> None:
    """Punto de entrada actual del proyecto.

    Ejecuta el Proceso 1 de la automatizacion:
    lectura de tickets OneTime Check, validacion y generacion de CSVs AP15.
    """
    process_all_tickets()


if __name__ == "__main__":
    main()
