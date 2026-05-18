import queue
import smtplib
import threading
import time
import tkinter as tk
import ctypes
from pathlib import Path
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText

import pythoncom

from src.common.config import get_settings
from src.common.logger import add_gui_callback
from src.common.settings_manager import SettingsManager
from src.common.system import kill_processes
from src.hda_web.client import HDAClient
from src.hda_web.ticket_processing import process_all_tickets
from src.mailer.client import SMTPMailClient
from src.sap.client import SAPGuiClient


def get_resource_path(relative_path: str) -> Path:
    return Path(__file__).resolve().parent / relative_path


class AutomationApp:
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("EssilorLuxottica - HDA Automation")
        self.root.geometry("920x780")
        self.root.minsize(860, 700)

        self.colors = {
            "bg": "#edf3f8",
            "navy": "#102945",
            "blue": "#1d5ea8",
            "cyan": "#00a7c4",
            "gold": "#d39d3c",
            "card": "#ffffff",
            "muted": "#627387",
            "border": "#d6e0ea",
            "log_bg": "#112032",
            "log_fg": "#edf6ff",
            "input_bg": "#f9fbfd",
        }

        self.root.configure(bg=self.colors["bg"])
        self.settings_manager = SettingsManager()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.is_running = False
        self.last_preflight_ok = False
        self.fields: dict[str, tk.StringVar] = {}
        self.field_widgets: list[tk.Widget] = []
        self.secret_entries: dict[str, tk.Entry] = {}
        self.secret_toggle_buttons: dict[str, tk.Button] = {}

        self._build_ui()
        self._load_saved_values()

        add_gui_callback(self.enqueue_log)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(200, self._flush_logs)

    def _build_ui(self) -> None:
        shell = tk.Frame(self.root, bg=self.colors["bg"])
        shell.pack(fill="both", expand=True)

        self.main_canvas = tk.Canvas(
            shell,
            bg=self.colors["bg"],
            highlightthickness=0,
            bd=0,
        )
        scrollbar = tk.Scrollbar(shell, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)

        outer = tk.Frame(self.main_canvas, bg=self.colors["bg"], padx=18, pady=18)
        self.main_canvas_window = self.main_canvas.create_window((0, 0), window=outer, anchor="nw")
        outer.bind("<Configure>", self._on_content_configure)
        self.main_canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_mousewheel(self.main_canvas)

        self._build_header(outer)

        content = tk.Frame(outer, bg=self.colors["bg"])
        content.pack(fill="both", expand=True, pady=(18, 0))

        form_card = self._make_card(content)
        form_card.pack(fill="x")

        tk.Label(
            form_card,
            text="Configuración rápida",
            bg=self.colors["card"],
            fg=self.colors["navy"],
            font=("Segoe UI", 15, "bold"),
        ).pack(anchor="w")
        tk.Label(
            form_card,
            text="Guarda credenciales y corre la automatización con una interfaz ligera y estable.",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 16))

        form = tk.Frame(form_card, bg=self.colors["card"])
        form.pack(fill="x")

        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)

        left_col = tk.Frame(form, bg=self.colors["card"])
        right_col = tk.Frame(form, bg=self.colors["card"])
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right_col.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self._build_form_column(
            left_col,
            [
                ("Usuario HDA", "hda_username", False),
                ("Contrasena HDA", "hda_password", True),
                ("Usuario SAP FMS", "sap_username_fms", False),
                ("Contrasena SAP FMS", "sap_password_fms", True),
                ("Sap connection name FMS", "sap_connection_name_fms", False),
                ("Usuario SAP AFS", "sap_username_afs", False),
                ("Contrasena SAP AFS", "sap_password_afs", True),
                ("Sap connection name AFS", "sap_connection_name_afs", False),
            ],
        )
        self._build_form_column(
            right_col,
            [
                ("Usuario correo", "smtp_username", False),
                ("Contrasena correo", "smtp_password", True),
                ("Correo Summary", "mail_summary_recipient", False),
                ("Correo FMS", "mail_fms_recipient", False),
                ("Correo AFS", "mail_afs_recipient", False),
            ],
        )

        button_row = tk.Frame(content, bg=self.colors["bg"])
        button_row.pack(fill="x", pady=(14, 12))

        self.save_button = self._make_button(
            button_row,
            text="Guardar",
            command=self.save_settings,
            bg=self.colors["card"],
            fg=self.colors["navy"],
            active_bg="#e7eef5",
        )
        self.save_button.pack(side="left")

        self.run_button = self._make_button(
            button_row,
            text="Correr Automatización",
            command=self.run_automation,
            bg=self.colors["blue"],
            fg="#ffffff",
            active_bg="#154d8a",
        )
        self.run_button.pack(side="left", padx=(8, 0))

        self.preflight_button = self._make_button(
            button_row,
            text="Verificar accesos",
            command=self.run_preflight,
            bg=self.colors["cyan"],
            fg="#ffffff",
            active_bg="#008ca4",
        )
        self.preflight_button.pack(side="left", padx=(8, 0))

        self.stop_button = self._make_button(
            button_row,
            text="Detener",
            command=self.request_stop,
            bg=self.colors["gold"],
            fg="#ffffff",
            active_bg="#b9882f",
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=(8, 0))

        log_card = self._make_card(content)
        log_card.pack(fill="both", expand=True)

        tk.Label(
            log_card,
            text="EssilorLuxottica Activity Log",
            bg=self.colors["card"],
            fg=self.colors["navy"],
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")
        tk.Label(
            log_card,
            text="Seguimiento en vivo de la ejecución de HDA, SAP y envío de reportes.",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 12))

        self.log_text = ScrolledText(
            log_card,
            height=24,
            wrap="word",
            font=("Consolas", 9),
            bg=self.colors["log_bg"],
            fg=self.colors["log_fg"],
            insertbackground=self.colors["log_fg"],
            relief="flat",
            bd=0,
            padx=12,
            pady=12,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def _on_content_configure(self, _event: tk.Event) -> None:
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self.main_canvas.itemconfigure(self.main_canvas_window, width=event.width)

    def _bind_mousewheel(self, widget: tk.Widget) -> None:
        widget.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event: tk.Event) -> None:
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_header(self, parent: tk.Widget) -> None:
        header = tk.Frame(parent, bg=self.colors["navy"], height=138, padx=22, pady=20)
        header.pack(fill="x")
        header.pack_propagate(False)

        accent = tk.Frame(header, bg=self.colors["cyan"], width=10)
        accent.pack(side="left", fill="y")

        content = tk.Frame(header, bg=self.colors["navy"], padx=18)
        content.pack(side="left", fill="both", expand=True)

        tk.Label(
            content,
            text="EssilorLuxottica",
            bg=self.colors["navy"],
            fg="#ffffff",
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            content,
            text="HDA Automation Console",
            bg=self.colors["navy"],
            fg="#d5e6f6",
            font=("Segoe UI", 13),
        ).pack(anchor="w", pady=(4, 0))

    def _make_card(self, parent: tk.Widget) -> tk.Frame:
        return tk.Frame(
            parent,
            bg=self.colors["card"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
            bd=0,
            padx=18,
            pady=18,
        )

    def _make_button(
        self,
        parent: tk.Widget,
        text: str,
        command,
        bg: str,
        fg: str,
        active_bg: str,
        state: str = "normal",
    ) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=fg,
            disabledforeground="#dce3ea",
            relief="flat",
            bd=0,
            padx=18,
            pady=10,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2" if state == "normal" else "arrow",
            state=state,
        )

    def _build_form_column(
        self,
        parent: tk.Widget,
        rows: list[tuple[str, str, bool]],
    ) -> None:
        parent.columnconfigure(0, weight=1)
        for index, (label, key, secret) in enumerate(rows):
            tk.Label(
                parent,
                text=label,
                bg=self.colors["card"],
                fg=self.colors["navy"],
                font=("Segoe UI", 10, "bold"),
            ).grid(row=index * 2, column=0, sticky="w", pady=(0, 6))

            variable = tk.StringVar()
            field_container = tk.Frame(parent, bg=self.colors["card"])
            field_container.grid(row=index * 2 + 1, column=0, sticky="ew", pady=(0, 12))
            field_container.columnconfigure(0, weight=1)

            entry = tk.Entry(
                field_container,
                textvariable=variable,
                width=32,
                show="*" if secret else "",
                relief="flat",
                bd=0,
                highlightthickness=1,
                highlightbackground=self.colors["border"],
                highlightcolor=self.colors["blue"],
                bg=self.colors["input_bg"],
                fg="#102033",
                insertbackground="#102033",
                font=("Segoe UI", 10),
            )
            entry.grid(row=0, column=0, sticky="ew", ipady=8)
            self.fields[key] = variable
            self.field_widgets.append(entry)

            if secret:
                toggle_button = tk.Button(
                    field_container,
                    text="Ver",
                    command=lambda field_key=key: self._toggle_secret_visibility(field_key),
                    relief="flat",
                    bd=0,
                    padx=12,
                    pady=8,
                    bg="#e7eef5",
                    fg=self.colors["navy"],
                    activebackground="#d8e4f0",
                    activeforeground=self.colors["navy"],
                    font=("Segoe UI", 9, "bold"),
                    cursor="hand2",
                )
                toggle_button.grid(row=0, column=1, padx=(8, 0))
                self.secret_entries[key] = entry
                self.secret_toggle_buttons[key] = toggle_button
                self.field_widgets.append(toggle_button)

    def _toggle_secret_visibility(self, key: str) -> None:
        entry = self.secret_entries.get(key)
        button = self.secret_toggle_buttons.get(key)
        if entry is None or button is None:
            return

        hidden = bool(entry.cget("show"))
        entry.configure(show="" if hidden else "*")
        button.configure(text="Ocultar" if hidden else "Ver")

    def _load_saved_values(self) -> None:
        settings = get_settings()
        self.fields["hda_username"].set(settings.hda_username)
        self.fields["hda_password"].set(settings.hda_password)
        self.fields["sap_username_fms"].set(settings.sap_username_fms)
        self.fields["sap_password_fms"].set(settings.sap_password_fms)
        self.fields["sap_connection_name_fms"].set(settings.sap_connection_name_fms)
        self.fields["sap_username_afs"].set(settings.sap_username_afs)
        self.fields["sap_password_afs"].set(settings.sap_password_afs)
        self.fields["sap_connection_name_afs"].set(settings.sap_connection_name_afs)
        self.fields["smtp_username"].set(settings.smtp_username)
        self.fields["smtp_password"].set(settings.smtp_password)
        self.fields["mail_summary_recipient"].set(settings.mail_summary_recipient)
        self.fields["mail_fms_recipient"].set(settings.mail_fms_recipient or settings.mail_usd_recipient)
        self.fields["mail_afs_recipient"].set(settings.mail_afs_recipient or settings.mail_cad_recipient)

    def _compose_settings_payload(self) -> dict:
        current = get_settings().to_dict()
        sap_username_fms = self.fields["sap_username_fms"].get().strip()
        sap_password_fms = self.fields["sap_password_fms"].get().strip()
        sap_connection_name_fms = self.fields["sap_connection_name_fms"].get().strip()
        sap_username_afs = self.fields["sap_username_afs"].get().strip()
        sap_password_afs = self.fields["sap_password_afs"].get().strip()
        sap_connection_name_afs = self.fields["sap_connection_name_afs"].get().strip()

        current.update(
            {
                "hda_username": self.fields["hda_username"].get().strip(),
                "hda_password": self.fields["hda_password"].get().strip(),
                "sap_username_fms": sap_username_fms,
                "sap_username_afs": sap_username_afs,
                "sap_password_fms": sap_password_fms,
                "sap_password_afs": sap_password_afs,
                "sap_connection_name_fms": sap_connection_name_fms,
                "sap_connection_name_afs": sap_connection_name_afs,
                "smtp_username": self.fields["smtp_username"].get().strip(),
                "smtp_password": self.fields["smtp_password"].get().strip(),
                "smtp_sender": self.fields["smtp_username"].get().strip(),
                "mail_summary_recipient": self.fields["mail_summary_recipient"].get().strip(),
                "mail_fms_recipient": self.fields["mail_fms_recipient"].get().strip(),
                "mail_afs_recipient": self.fields["mail_afs_recipient"].get().strip(),
                "mail_test_recipient": "",
                "mail_usd_recipient": self.fields["mail_fms_recipient"].get().strip(),
                "mail_cad_recipient": self.fields["mail_afs_recipient"].get().strip(),
            }
        )
        return current

    def save_settings(self, notify: bool = True) -> bool:
        try:
            payload = self._compose_settings_payload()
            self.settings_manager.save_settings(payload)
            if notify:
                messagebox.showinfo("Guardado", "La configuración quedó guardada.")
            return True
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo guardar la configuración.\n\n{exc}")
            return False

    def run_automation(self) -> None:
        if self.is_running:
            return

        if not self.save_settings(notify=False):
            return

        self.is_running = True
        self.cancel_event.clear()
        self._set_running_state(True)
        self.enqueue_log("Iniciando automatizacion...")

        self.worker_thread = threading.Thread(
            target=self._worker_main,
            name="automation-worker",
            daemon=True,
        )
        self.worker_thread.start()

    def _worker_main(self) -> None:
        pythoncom.CoInitialize()
        self._prevent_sleep()
        try:
            process_all_tickets(abort_event=self.cancel_event)
            if self.cancel_event.is_set():
                self.enqueue_log("Proceso cancelado por el usuario.")
            else:
                self.enqueue_log("Proceso finalizado exitosamente.")
        except Exception as exc:
            self.enqueue_log(f"ERROR: {exc}")
        finally:
            self._allow_sleep()
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self._set_running_state(False))

    def request_stop(self) -> None:
        if not self.is_running:
            return
        should_stop = messagebox.askyesno(
            "Confirmar",
            "¿Estás seguro de detener el proceso?\n\nEsto cerrará instantáneamente el navegador, SAP y Excel."
        )
        if not should_stop:
            return
            
        self.cancel_event.set()
        self.enqueue_log("Se solicitó detener la ejecución de forma forzada. Abortando procesos...")
        
        try:
            kill_processes(["msedge.exe", "msedgedriver.exe", "saplogon.exe", "sapgui.exe", "excel.exe"])
        except Exception as e:
            self.enqueue_log(f"Fallo al intentar matar procesos: {e}")

    def run_preflight(self) -> None:
        if self.is_running:
            return
        if not self.save_settings(notify=False):
            return
        self.is_running = True
        self.cancel_event.clear()
        self._set_running_state(True)
        self.worker_thread = threading.Thread(
            target=self._preflight_only_worker,
            name="preflight-worker",
            daemon=True,
        )
        self.worker_thread.start()

    def _preflight_only_worker(self) -> None:
        pythoncom.CoInitialize()
        try:
            ok, message = self._run_preflight_checks()
            if ok:
                self.last_preflight_ok = True
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Validacion correcta",
                        "HDA, SAP y SMTP fueron validados correctamente.",
                    ),
                )
            else:
                self.last_preflight_ok = False
                self.enqueue_log(f"ERROR preflight: {message}")
                self.root.after(
                    0,
                    lambda m=message: messagebox.showerror(
                        "Validacion fallida",
                        f"No se pudo validar uno de los accesos.\n\n{m}\n\nLa automatizacion no se iniciara.",
                    ),
                )
        finally:
            pythoncom.CoUninitialize()
            self.root.after(0, lambda: self._set_running_state(False))

    def _run_preflight_checks(self) -> tuple[bool, str]:
        self.enqueue_log("Verificando accesos de HDA, SAP y SMTP...")
        try:
            self._check_hda_access()
            self.enqueue_log("OK | HDA login correcto.")

            self._check_sap_access("FMS")
            self.enqueue_log("OK | SAP FMS login correcto.")

            self._check_sap_access("AFS")
            self.enqueue_log("OK | SAP AFS login correcto.")

            self._check_smtp_access()
            self.enqueue_log("OK | SMTP autenticado correctamente.")
            return True, ""
        except Exception as exc:
            return False, self._format_preflight_error(exc)

    def _check_hda_access(self) -> None:
        client = HDAClient(get_settings())
        try:
            client.start()
            client.login()
            if client.is_login_screen_visible():
                raise RuntimeError("HDA rechazó las credenciales o no avanzó después del login.")
            client.click_payments_tile()
        finally:
            client.close()

    def _check_sap_access(self, system_group: str) -> None:
        sap = SAPGuiClient(get_settings())
        try:
            tcode = sap.verify_access(system_group)
            self.enqueue_log(f"OK | SAP {system_group} listo y navegó a {tcode}.")
        finally:
            sap.close()

    def _check_smtp_access(self) -> None:
        try:
            SMTPMailClient(get_settings()).test_connection()
        except smtplib.SMTPAuthenticationError as exc:
            raise RuntimeError(
                "Correo/SMTP rechazo las credenciales. Revisa Usuario correo y Contrasena correo."
            ) from exc
        except smtplib.SMTPException as exc:
            raise RuntimeError(
                f"Correo/SMTP no pudo autenticarse correctamente: {exc}"
            ) from exc
        except OSError as exc:
            raise RuntimeError(
                f"Correo/SMTP no pudo conectarse al servidor configurado: {exc}"
            ) from exc

    def _format_preflight_error(self, exc: Exception) -> str:
        message = str(exc).strip()
        if not message:
            return "La validacion fallo por un error no especificado."

        normalized = message.casefold()
        if "authentication unsuccessful" in normalized or "535 5.7.139" in normalized:
            return "Correo/SMTP rechazo las credenciales. Revisa Usuario correo y Contrasena correo."
        if "prod.outlook.com" in normalized and "credential" in normalized:
            return "Correo/SMTP rechazo la autenticacion. Revisa Usuario correo y Contrasena correo."
        return message

    def _prevent_sleep(self) -> None:
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(
                self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED | self.ES_DISPLAY_REQUIRED
            )
            self.enqueue_log("Modo activo: Windows no entrara en suspension durante la ejecucion.")
        except Exception as exc:
            self.enqueue_log(f"ADVERTENCIA: No se pudo bloquear la suspension del equipo: {exc}")

    def _allow_sleep(self) -> None:
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(self.ES_CONTINUOUS)
            self.enqueue_log("Modo activo liberado: Windows puede volver a usar su configuracion normal de energia.")
        except Exception as exc:
            self.enqueue_log(f"ADVERTENCIA: No se pudo restaurar el estado de energia normal: {exc}")

    def enqueue_log(self, message: str) -> None:
        if message:
            self.log_queue.put(str(message))

    def _flush_logs(self) -> None:
        while True:
            try:
                line = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"{line}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.root.after(200, self._flush_logs)

    def _set_running_state(self, running: bool) -> None:
        self.is_running = running
        for widget in self.field_widgets:
            widget.configure(state="disabled" if running else "normal")

        self.run_button.configure(
            state="disabled" if running else "normal",
            bg="#8eb3df" if running else self.colors["blue"],
            cursor="arrow" if running else "hand2",
        )
        self.preflight_button.configure(
            state="disabled" if running else "normal",
            bg="#7fd0dc" if running else self.colors["cyan"],
            cursor="arrow" if running else "hand2",
        )
        self.stop_button.configure(
            state="normal" if running else "disabled",
            bg=self.colors["gold"] if running else "#c8d0da",
            cursor="hand2" if running else "arrow",
        )
        self.save_button.configure(
            state="disabled" if running else "normal",
            bg="#edf2f7" if running else self.colors["card"],
            cursor="arrow" if running else "hand2",
        )

    def _on_close(self) -> None:
        if self.is_running:
            should_close = messagebox.askyesno(
                "Cerrar",
                "Hay una ejecución en curso. ¿Quieres cerrar la aplicación y cancelar el proceso?",
            )
            if not should_close:
                return
            self.cancel_event.set()
            self.enqueue_log("Cerrando aplicación. Se canceló la ejecución activa.")
            worker = self.worker_thread
            if worker and worker.is_alive():
                worker.join(timeout=8)
            kill_processes(["msedgedriver.exe"])
            time.sleep(1)

        self.root.destroy()


def start_gui() -> None:
    root = tk.Tk()
    icon_path = get_resource_path("src/gui/icon.png")
    if icon_path.exists():
        icon_image = tk.PhotoImage(file=str(icon_path))
        root.iconphoto(True, icon_image)
        root._icon_image_ref = icon_image
    AutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    start_gui()
