from __future__ import annotations

import mimetypes
import smtplib
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path

from src.common.config import Settings


@dataclass
class MailSendResult:
    recipients: list[str]
    bcc: list[str]
    subject: str
    attachments: list[str]


class MailClient:
    """Cliente base de correo."""

    def send_message(
        self,
        subject: str,
        body: str,
        attachments: list[str],
        recipients: list[str] | None = None,
        bcc: list[str] | None = None,
        html_body: str | None = None,
    ) -> MailSendResult:
        raise NotImplementedError

    def send_process_report(
        self,
        subject: str,
        body: str,
        attachments: list[str],
    ) -> MailSendResult:
        raise NotImplementedError

    def fetch_bot_responses(self) -> list:
        raise NotImplementedError


class SMTPMailClient(MailClient):
    """Cliente SMTP para envio del resumen del Proceso 1."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def send_message(
        self,
        subject: str,
        body: str,
        attachments: list[str],
        recipients: list[str] | None = None,
        bcc: list[str] | None = None,
        html_body: str | None = None,
    ) -> MailSendResult:
        recipients = self._resolve_recipients(recipients)
        bcc = self._resolve_bcc(bcc)
        if not recipients:
            raise ValueError("No email recipients configured.")

        sender = self.settings.smtp_sender or self.settings.smtp_username
        if not sender:
            raise ValueError("No SMTP sender configured.")

        message = EmailMessage()
        message["From"] = sender
        message["To"] = ", ".join(recipients)
        if bcc:
            message["Bcc"] = ", ".join(bcc)
        message["Subject"] = subject
        message.set_content(body)
        if html_body:
            message.add_alternative(html_body, subtype="html")

        for attachment in attachments:
            attachment_path = Path(attachment)
            if not attachment_path.exists():
                continue
            mime_type, _ = mimetypes.guess_type(attachment_path.name)
            if mime_type:
                maintype, subtype = mime_type.split("/", 1)
            else:
                maintype, subtype = "application", "octet-stream"
            with attachment_path.open("rb") as attachment_file:
                message.add_attachment(
                    attachment_file.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=attachment_path.name,
                )

        all_recipients = recipients + bcc
        with self._connect() as smtp:
            supports_auth = "auth" in (smtp.esmtp_features or {})
            if (
                self.settings.smtp_username
                and self.settings.smtp_password
                and supports_auth
            ):
                smtp.login(self.settings.smtp_username, self.settings.smtp_password)
            smtp.send_message(message, to_addrs=all_recipients)

        return MailSendResult(
            recipients=recipients,
            bcc=bcc,
            subject=subject,
            attachments=attachments,
        )

    def send_process_report(
        self,
        subject: str,
        body: str,
        attachments: list[str],
    ) -> MailSendResult:
        return self.send_message(subject=subject, body=body, attachments=attachments)

    def fetch_bot_responses(self) -> list:
        raise NotImplementedError

    def _connect(self) -> smtplib.SMTP:
        host = self.settings.smtp_host
        port = self.settings.smtp_port
        if not host or not port:
            raise ValueError("SMTP host/port are not configured.")

        if self.settings.smtp_use_ssl:
            return smtplib.SMTP_SSL(host, port, timeout=30)

        smtp = smtplib.SMTP(host, port, timeout=30)
        smtp.ehlo()
        if self.settings.smtp_use_tls:
            smtp.starttls()
            smtp.ehlo()
        return smtp

    def _resolve_recipients(self, recipients: list[str] | None = None) -> list[str]:
        if self.settings.mail_test_recipient:
            return self._split_recipients(self.settings.mail_test_recipient)

        if recipients is not None:
            return [recipient for recipient in recipients if recipient]

        default_recipients = [
            self.settings.mail_primary_recipient,
            self.settings.mail_secondary_recipient,
        ]
        combined: list[str] = []
        for recipient in default_recipients:
            combined.extend(self._split_recipients(recipient))
        return combined

    def _resolve_bcc(self, bcc: list[str] | None = None) -> list[str]:
        if bcc is not None:
            return [recipient for recipient in bcc if recipient]
        return self._split_recipients(self.settings.mail_bcc_recipient)

    @staticmethod
    def _split_recipients(raw_value: str) -> list[str]:
        if not raw_value:
            return []
        normalized = raw_value.replace(";", ",")
        return [part.strip() for part in normalized.split(",") if part.strip()]
