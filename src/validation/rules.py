from src.common.models import TicketRecord


class ValidationEngine:
    """Motor de reglas de negocio.

    La regla exacta de suspensión sigue en levantamiento.
    """

    def validate(self, ticket: TicketRecord) -> TicketRecord:
        return ticket
