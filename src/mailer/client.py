class MailClient:
    """Cliente base de correo.

    Pendiente definir transporte final y logica de jueves.
    """

    def send_excel(self, excel_path: str, recipient: str) -> None:
        raise NotImplementedError

    def fetch_bot_responses(self) -> list:
        raise NotImplementedError
