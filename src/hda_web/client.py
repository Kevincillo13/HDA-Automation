from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Any

from selenium import webdriver
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

from src.common.config import Settings, get_settings
from src.common.logger import get_logger
from src.common.models import TicketRecord
from src.common.run_context import get_run_id


class HDAClient:
    """Cliente inicial para HDA usando Selenium con Edge."""

    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()
        self.logger = get_logger("hda_client")
        self.driver: webdriver.Edge | None = None

    def start(self) -> None:
        """Inicia Edge y abre la pagina de login."""
        os.makedirs(self.settings.download_dir, exist_ok=True)
        if self.settings.evidence_enabled:
            os.makedirs(self.settings.evidence_dir, exist_ok=True)

        options = Options()
        options.use_chromium = True
        options.binary_location = self.settings.browser_binary_path
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(Path(self.settings.download_dir).resolve()),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True,
            },
        )
        options.set_capability("goog:loggingPrefs", {"browser": "ALL", "performance": "ALL"})
        if self.settings.browser_headless:
            options.add_argument("--headless=new")

        self.driver = webdriver.Edge(options=options, service=Service())
        self.driver.get(self.settings.hda_url)
        self.logger.info("HDA login page opened.")
        self.logger.info("Browser binary: %s", self.settings.browser_binary_path)
        self.logger.info("Initial URL: %s", self.driver.current_url)
        self._pause()

    def login(self) -> None:
        """Hace login con usuario y contraseña."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if not self.settings.hda_username or not self.settings.hda_password:
            raise ValueError("Missing HDA credentials in local environment.")

        wait = WebDriverWait(self.driver, 30)
        self.logger.info("Waiting for login form fields.")
        wait.until(ec.presence_of_element_located((By.CSS_SELECTOR, "#txtUsername-inputEl")))
        self.logger.info("Login form detected.")

        self.driver.find_element(By.CSS_SELECTOR, "#txtUsername-inputEl").send_keys(
            self.settings.hda_username
        )
        self.logger.info("Username entered.")
        self.driver.find_element(By.CSS_SELECTOR, "#txtPassword-inputEl").send_keys(
            self.settings.hda_password
        )
        self.logger.info("Password entered.")
        self._pause()
        self.driver.find_element(By.CSS_SELECTOR, "#cmdLogin-btnEl").click()
        self.logger.info("Login button clicked.")

        self._pause(2.0)
        try:
            wait.until_not(ec.presence_of_element_located((By.CSS_SELECTOR, "#txtUsername-inputEl")))
        except TimeoutException:
            self.logger.warning("Login form still visible after submit timeout.")

        self.logger.info("Login submitted. Current URL: %s", self.driver.current_url)
        self.logger.info("Current title: %s", self.driver.title)

    def click_payments_tile(self) -> None:
        """Espera el tile Payments y hace clic."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        wait = WebDriverWait(self.driver, 60)
        tile_selector = 'div.tile[action="Payment"]'

        self.logger.info("Waiting for Payments tile to render.")
        wait.until(ec.presence_of_element_located((By.CSS_SELECTOR, tile_selector)))
        tiles = self.driver.find_elements(By.CSS_SELECTOR, tile_selector)
        self.logger.info("Payments tiles found: %s", len(tiles))

        payment_tile = wait.until(
            ec.element_to_be_clickable((By.CSS_SELECTOR, tile_selector))
        )
        tile_text = payment_tile.text.replace("\n", " ").strip()
        self.logger.info("Clicking Payments tile with text: %s", tile_text)
        self._scroll_into_view(payment_tile)
        self._pause()
        payment_tile.click()
        self.logger.info("Payments tile clicked.")
        self._pause(2.0)

        self._ensure_payments_tab_active()
        self.log_debug_state("after_payments_tile_click")

    def read_payment_grid_rows(self) -> list[TicketRecord]:
        """Lee las filas del grid de Payments sin depender del viewport."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        self._ensure_payments_tab_active()
        wait = WebDriverWait(self.driver, 60)
        grid_selector = "div.x-grid-view"

        self.logger.info("Waiting for Payments grid rows.")
        wait.until(ec.presence_of_element_located((By.CSS_SELECTOR, grid_selector)))
        self._pause(1.0)

        scroll_state = self._get_grid_scroll_state()
        self.logger.info(
            "Payments grid discovery | active_payments_tab=%s | current_url=%s | scroll_state=%s",
            self._is_payments_tab_active(),
            self.driver.current_url,
            scroll_state,
        )

        discovered_records: dict[str, TicketRecord] = {}
        max_pages = 20
        max_scroll_passes = 12

        for page_index in range(1, max_pages + 1):
            self.logger.info("Scanning Payments grid page %s", page_index)
            page_records: dict[str, TicketRecord] = {}

            for pass_index in range(1, max_scroll_passes + 1):
                snapshot = self._collect_grid_records_from_dom()
                self.logger.info(
                    "Grid page %s scan pass %s | dom_rows=%s | rows_with_ticket_id=%s",
                    page_index,
                    pass_index,
                    len(snapshot),
                    len([row for row in snapshot if row.ticket_id]),
                )

                for record in snapshot:
                    if not record.ticket_id:
                        continue
                    page_records.setdefault(record.ticket_id, record)
                    discovered_records.setdefault(record.ticket_id, record)

                moved = self._scroll_grid_container()
                self.logger.info(
                    "Grid page %s scan pass %s | page_unique_ticket_ids=%s | total_unique_ticket_ids=%s | scroll_moved=%s",
                    page_index,
                    pass_index,
                    len(page_records),
                    len(discovered_records),
                    moved,
                )
                if not moved:
                    break
                self._pause(0.5)

            page_ticket_ids = [ticket_id for ticket_id in page_records if ticket_id]
            self.logger.info(
                "Completed Payments grid page %s | unique_ticket_ids=%s | first_ticket=%s | last_ticket=%s",
                page_index,
                len(page_ticket_ids),
                page_ticket_ids[0] if page_ticket_ids else "<none>",
                page_ticket_ids[-1] if page_ticket_ids else "<none>",
            )

            if not self._go_to_next_grid_page(page_ticket_ids, page_index):
                break

        ticket_records = list(discovered_records.values())
        self.logger.info("Filtered grid rows with ticket IDs: %s", len(ticket_records))
        if not ticket_records:
            self._log_grid_debug()
        for index, record in enumerate(ticket_records, start=1):
            self.logger.info(
                "Grid row %s | ticket_id=%s | created=%s | payment_method=%s | company=%s | type=%s | status=%s | subject=%s",
                index,
                record.ticket_id,
                record.created,
                record.payment_method,
                record.company,
                record.ticket_type,
                record.hda_status,
                record.subject,
            )

        otc_records = [
            ticket for ticket in ticket_records 
            if ticket.payment_method.strip() == "OneTime Check"
            and ticket.hda_status.strip() in ["Open", "Assigned"]
        ]
        self.logger.info("OneTime Check candidates visible: %s", len(otc_records))
        for ticket in otc_records:
            self.logger.info(
                "OneTime Check candidate | ticket_id=%s | company=%s | subject=%s",
                ticket.ticket_id,
                ticket.company,
                ticket.subject,
            )

        return ticket_records

    def open_ticket_by_id(self, ticket_id: str) -> None:
        """Abre un ticket desde el grid con varias estrategias."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if self._is_ticket_detail_open(ticket_id):
            self.logger.info("Ticket already open in detail view: %s", ticket_id)
            return

        self._ensure_payments_tab_active()
        wait = WebDriverWait(self.driver, 60)
        self.logger.info("Waiting for ticket row to open: %s", ticket_id)
        ticket_cell = wait.until(
            ec.presence_of_element_located(
                (By.XPATH, f"//div[contains(@class,'ticket-id') and normalize-space()='{ticket_id}']")
            )
        )
        row = ticket_cell.find_element(
            By.XPATH,
            "./ancestor::table[contains(@class,'x-grid-item')]",
        )
        self._scroll_into_view(ticket_cell)
        self._pause()

        strategies = [
            ("context menu open option", lambda: self._open_ticket_from_context_menu(row, ticket_id)),
            ("double click ticket cell", lambda: ActionChains(self.driver).double_click(ticket_cell).perform()),
            ("double click row", lambda: ActionChains(self.driver).double_click(row).perform()),
            ("javascript double click ticket cell", lambda: self.driver.execute_script(
                """
                const target = arguments[0];
                target.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));
                """,
                ticket_cell,
            )),
            ("single click row and press enter", lambda: self._open_ticket_with_enter(row)),
        ]

        for label, action in strategies:
            try:
                self.logger.info("Trying ticket open strategy [%s] for %s", label, ticket_id)
                self._ensure_payments_tab_active()
                self._scroll_into_view(ticket_cell)
                self._pause()
                action()
                self._pause(2.0)
                if self._wait_for_ticket_detail(ticket_id, timeout=12):
                    self.logger.info("Ticket opened with strategy [%s]: %s", label, ticket_id)
                    return
                self.logger.info("Strategy [%s] did not open ticket %s", label, ticket_id)
            except Exception as exc:
                self.logger.warning(
                    "Strategy [%s] failed for ticket %s: %s",
                    label,
                    ticket_id,
                    exc,
                )

        raise RuntimeError(f"Unable to open ticket {ticket_id} with available strategies.")

    def close_active_ticket_tab(self) -> None:
        """
        Cierra la pestaña del ticket que está activa y espera a que la de Payments
        vuelva a ser la principal.
        """
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        self.logger.info("Closing active ticket tab.")
        try:
            # Selector for an active tab that is NOT the 'Payments' tab.
            active_tab_selector = (
                By.XPATH,
                "//a[contains(@class, 'x-tab-active') and not(.//span[normalize-space()='Payments'])]",
            )
            active_tab = WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located(active_tab_selector)
            )

            # Find the close button within that specific tab
            close_button = active_tab.find_element(By.CSS_SELECTOR, "span.x-tab-close-btn")
            self._scroll_into_view(close_button)
            self.driver.execute_script("arguments[0].click();", close_button)
            self.logger.info("Active ticket tab close button clicked.")

            # Wait until the 'Payments' tab is the active one again
            WebDriverWait(self.driver, 15).until(
                lambda _: self._is_payments_tab_active()
            )
            self.logger.info("Successfully returned to Payments tab.")
            self._pause(1.0)
        except TimeoutException:
            self.logger.warning("Could not find an active ticket tab to close, or failed to return to Payments tab.")
            self.log_debug_state("close_tab_timeout")
            self.take_screenshot("close_tab_timeout")
            # If closing fails, try to force-navigate back to the main grid
            self._ensure_payments_tab_active()
        except Exception as exc:
            self.logger.error("An unexpected error occurred while closing ticket tab: %s", exc)
            self.log_debug_state("close_tab_error")
            self.take_screenshot("close_tab_error")
            raise

    def suspend_ticket(self, ticket_id: str, reasons: list[str]) -> None:
        """
        Suspende un ticket usando la UI de HDA.
        """
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        wait = WebDriverWait(self.driver, 20)
        self.logger.info("Iniciando proceso de suspensión en HDA para el ticket: %s", ticket_id)

        try:
            # 1. Clic en Autoasignar
            self.logger.info("Buscando botón 'Autoasignar' (basado en icono)...")
            autoasignar_xpath = "//a[.//span[contains(@class, 'icon-ownership')]]"
            
            start_time = time.time()
            autoasignar_clicked = False
            while time.time() - start_time < 5:
                botones = self.driver.find_elements(By.XPATH, autoasignar_xpath)
                visible = next((b for b in botones if b.is_displayed()), None)
                if visible:
                    self.driver.execute_script("arguments[0].click();", visible)
                    autoasignar_clicked = True
                    break
                time.sleep(0.5)

            if autoasignar_clicked:
                self._pause(1.0)

                # 2. Aceptar el Autoasignar (Sí)
                self.logger.info("Confirmando 'Autoasignar' (Buscando botón Sí/OK/Yes)...")
                
                # Buscamos cualquier botón visible que tenga el texto de confirmación
                si_xpath = "//a[contains(@class, 'x-btn')]//span[normalize-space(text())='Sí' or normalize-space(text())='Yes' or normalize-space(text())='OK' or normalize-space(text())='Aceptar']"
                
                start_confirm = time.time()
                si_clicked = False
                while time.time() - start_confirm < 10:
                    btns = self.driver.find_elements(By.XPATH, si_xpath)
                    # Filtramos por el que realmente se ve
                    visible_si = next((b for b in btns if b.is_displayed()), None)
                    if visible_si:
                        self.driver.execute_script("arguments[0].click();", visible_si)
                        si_clicked = True
                        break
                    time.sleep(0.5)
                
                if not si_clicked:
                    raise TimeoutException("No se encontró el botón de confirmación 'Sí' visible.")
                
                self._pause(2.0)
                self.logger.info("Ticket autoasignado con éxito.")
            else:
                self.logger.info("El botón 'Autoasignar' no está disponible o ya está asignado. Continuando...")

            # --- PASO 2: Escribir Solución ---
            self.logger.info("Escribiendo solución...")
            solution_text = "SUSPENDED\n\nReasons:\n- " + "\n- ".join(reasons)
            
            self._pause(2.0)  # Pausa extra para que el editor TinyMCE se estabilice tras autoasignar
            
            try:
                start_sol = time.time()
                written = False
                while time.time() - start_sol < 15:
                    try:
                        iframes = self.driver.find_elements(By.CSS_SELECTOR, "iframe[id$='_SolutionHTML_SolutionHTML_ifr']")
                        visible_iframe = next((f for f in iframes if f.is_displayed()), None)
                        
                        if visible_iframe:
                            self.driver.switch_to.frame(visible_iframe)
                            # Esperamos a que el body esté presente y sea interactuable
                            body = WebDriverWait(self.driver, 5).until(
                                ec.element_to_be_clickable((By.TAG_NAME, "body"))
                            )
                            body.click()
                            self._pause(0.5)
                            
                            # Limpieza y escritura
                            self.driver.execute_script("document.body.innerHTML = '';")
                            body.send_keys(solution_text)
                            
                            from selenium.webdriver.common.keys import Keys
                            body.send_keys(Keys.TAB)
                            
                            self.driver.switch_to.default_content()
                            written = True
                            self.logger.info("Solución escrita correctamente.")
                            break
                        else:
                            self.logger.info("Esperando a que el cuadro de solución sea visible...")
                    except Exception as e:
                        self.logger.warning(f"Reintentando escritura de solución por error temporal: {e}")
                        self.driver.switch_to.default_content()
                    
                    time.sleep(1.0)
                
                if not written:
                    self.logger.warning("No se pudo escribir la solución tras varios intentos.")
                    
            except Exception as e:
                self.logger.error("Error crítico al escribir la solución: %s", e)
                self.driver.switch_to.default_content()

            self.logger.info("--- PASO 2 COMPLETADO: Solución escrita ---")

            # --- PASO 3: Seleccionar Categoría ---
            self.logger.info("Abriendo menú de Categoría...")
            try:
                # Localizamos el selector de categoría
                cat_arrow = WebDriverWait(self.driver, 10).until(
                    ec.element_to_be_clickable((By.CSS_SELECTOR, "button[id$='_TicketCategoryID-trigger-picker']"))
                )
                self.driver.execute_script("arguments[0].click();", cat_arrow)
                self._pause(1.5)

                # Buscamos y expandimos 'Payment Request'
                try:
                    self.logger.info("Buscando y expandiendo 'Payment Request'...")
                    # Buscamos el texto en el árbol (ignorando mayúsculas/minúsculas con translate si fuera necesario, pero contains es suficiente aquí)
                    payment_node = WebDriverWait(self.driver, 5).until(
                        ec.presence_of_element_located((By.XPATH, "//*[contains(@class, 'x-tree-node-text') and contains(normalize-space(.), 'Payment Request')]"))
                    )
                    self._scroll_into_view(payment_node)
                    
                    # Verificamos si ya está expandido
                    parent_row = payment_node.find_element(By.XPATH, "./ancestor::tr")
                    is_expanded = "x-grid-tree-node-expanded" in parent_row.get_attribute("class")
                    
                    if not is_expanded:
                        # Intentamos darle clic al expander o al texto mismo
                        try:
                            # Buscamos cualquier elemento de expansión (flecha, icono, etc)
                            expander = parent_row.find_element(By.XPATH, ".//*[contains(@class, 'x-tree-expander') or contains(@class, 'x-tree-elbow-img')]")
                            self.driver.execute_script("arguments[0].click();", expander)
                        except Exception:
                            # Si falla el expander, clic al texto
                            self.driver.execute_script("arguments[0].click();", payment_node)
                        
                        self._pause(1.5) # Esperamos a que se despliegue la lista
                    else:
                        self.logger.info("'Payment Request' ya está expandido.")
                except Exception as e:
                    self.logger.info(f"Nota: Problema al expandir 'Payment Request': {e}")

                # Seleccionamos 'Non-AP15 One-time'
                self.logger.info("Seleccionando 'Non-AP15 One-time'...")
                non_ap15_xpath = "//*[contains(@class, 'x-tree-node-text') and contains(normalize-space(.), 'Non-AP15')]"
                non_ap15_text = WebDriverWait(self.driver, 10).until(
                    ec.presence_of_element_located((By.XPATH, non_ap15_xpath))
                )
                self._scroll_into_view(non_ap15_text)
                
                # Buscamos el checkbox en esa misma fila y le damos clic
                checkbox = non_ap15_text.find_element(By.XPATH, "./ancestor::tr//input")
                self.driver.execute_script("arguments[0].click();", checkbox)
                
                self.logger.info("Categoría seleccionada con éxito.")
                
            except Exception as e:
                self.logger.error("Error al seleccionar la categoría: %s", e)

            # --- PASO 4: Clic en 'Cambiar estado...' ---
            self.logger.info("Haciendo clic en 'Cambiar estado...' (basado en icono)...")
            cambiar_estado_xpath = "//a[.//span[contains(@class, 'icon-changestatus')]]"
            
            try:
                start_ce = time.time()
                ce_clicked = False
                while time.time() - start_ce < 10:
                    botones = self.driver.find_elements(By.XPATH, cambiar_estado_xpath)
                    # Filtramos por visibilidad
                    visible_ce = next((b for b in botones if b.is_displayed()), None)
                    if visible_ce:
                        self.driver.execute_script("arguments[0].click();", visible_ce)
                        ce_clicked = True
                        break
                    time.sleep(0.5)
                
                if not ce_clicked:
                    raise TimeoutException("No se encontró el botón de 'Cambiar estado...' visible.")
                
                self.logger.info("Botón 'Cambiar estado...' clicado con éxito. Esperando alerta...")
                self._pause(1.5)

                # --- PASO 5: Aceptar alerta de confirmación (Sí) ---
                self.logger.info("Confirmando alerta de cambio de estado (Buscando botón Sí/OK/Yes)...")
                si_xpath = "//a[contains(@class, 'x-btn')]//span[normalize-space(text())='Sí' or normalize-space(text())='Yes' or normalize-space(text())='OK' or normalize-space(text())='Aceptar']"
                
                start_confirm = time.time()
                confirm_clicked = False
                while time.time() - start_confirm < 10:
                    btns = self.driver.find_elements(By.XPATH, si_xpath)
                    visible_si = next((b for b in btns if b.is_displayed()), None)
                    if visible_si:
                        self.driver.execute_script("arguments[0].click();", visible_si)
                        confirm_clicked = True
                        break
                    time.sleep(0.5)
                
                if confirm_clicked:
                    self.logger.info("Alerta confirmada con éxito.")
                else:
                    self.logger.warning("No se encontró la alerta de confirmación o se cerró sola.")

                self._pause(2.0)

                # --- PASO 6: Seleccionar 'Suspended' y Ejecutar ---
                self.logger.info("Abriendo menú de nuevo estado...")
                status_arrow = WebDriverWait(self.driver, 10).until(
                    ec.element_to_be_clickable((By.CSS_SELECTOR, "button[id$='_cbStatus-trigger-picker']"))
                )
                self.driver.execute_script("arguments[0].click();", status_arrow)
                self._pause(1.0)

                self.logger.info("Buscando opción 'Suspended'...")
                # Buscamos en los menús desplegables (boundlist)
                suspended_xpath = "//div[contains(@class, 'x-boundlist')]//li[contains(normalize-space(.), 'Suspend')]"
                suspended_option = WebDriverWait(self.driver, 10).until(
                    ec.element_to_be_clickable((By.XPATH, suspended_xpath))
                )
                self.driver.execute_script("arguments[0].click();", suspended_option)
                self._pause(1.0)

                self.logger.info("Haciendo clic en 'Ejecutar'...")
                # El botón de guardado final suele tener un ID que termina en _BtnOK
                start_exec = time.time()
                exec_clicked = False
                while time.time() - start_exec < 10:
                    btns = self.driver.find_elements(By.CSS_SELECTOR, "a[id$='_BtnOK']")
                    visible_btn = next((b for b in btns if b.is_displayed()), None)
                    if visible_btn:
                        self.driver.execute_script("arguments[0].click();", visible_btn)
                        exec_clicked = True
                        break
                    time.sleep(0.5)
                
                if exec_clicked:
                    self.logger.info("¡Proceso de suspensión completado con éxito!")
                else:
                    raise TimeoutException("No se encontró el botón 'Ejecutar' visible.")

                self._pause(3.0)
                
            except Exception as e:
                self.logger.error("Error en el paso final de cambio de estado: %s", e)

            self.logger.info("--- FLUJO COMPLETADO ---")

        except Exception as exc:
            self.logger.error("Error al suspender el ticket %s: %s", ticket_id, exc)
            self.take_screenshot(f"error_suspend_{ticket_id}")
            raise

    def take_screenshot(self, name: str) -> Path:
        """Guarda screenshot en la carpeta de evidencia."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if not self.settings.evidence_enabled:
            self.logger.info("Skipping screenshot [%s] because evidence capture is disabled.", name)
            return Path(self.settings.evidence_dir) / f"{get_run_id()}_{name}.png"

        safe_name = f"{name}.png"
        path = Path(self.settings.evidence_dir) / f"{get_run_id()}_{safe_name}"
        self.driver.save_screenshot(str(path))
        self.logger.info("Screenshot saved: %s", path)
        return path

    def save_page_source(self, name: str) -> Path:
        """Guarda el HTML actual para depuracion."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if not self.settings.evidence_enabled:
            self.logger.info("Skipping page source save [%s] because evidence capture is disabled.", name)
            return Path(self.settings.evidence_dir) / f"{get_run_id()}_{name}.html"

        safe_name = f"{name}.html"
        path = Path(self.settings.evidence_dir) / f"{get_run_id()}_{safe_name}"
        path.write_text(self.driver.page_source, encoding="utf-8")
        self.logger.info("Page source saved: %s", path)
        return path

    def log_debug_state(self, label: str) -> dict[str, Any]:
        """Escribe al log el estado visible del navegador."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        try:
            window_rect = self.driver.get_window_rect()
        except InvalidSessionIdException:
            self.logger.warning("Browser session already closed while collecting debug state [%s].", label)
            raise
        except Exception as exc:
            window_rect = {"error": str(exc)}

        state = {
            "label": label,
            "url": self.driver.current_url,
            "title": self.driver.title,
            "ready_state": self._get_ready_state(),
            "login_screen_visible": self.is_login_screen_visible(),
            "username_fields": len(self.driver.find_elements(By.CSS_SELECTOR, "#txtUsername-inputEl")),
            "password_fields": len(self.driver.find_elements(By.CSS_SELECTOR, "#txtPassword-inputEl")),
            "login_buttons": len(self.driver.find_elements(By.CSS_SELECTOR, "#cmdLogin-btnEl")),
            "page_tabs": len(self.driver.window_handles),
            "active_payments_tab": self._is_payments_tab_active(),
            "window_rect": window_rect,
        }
        self.logger.info("Browser state [%s]: %s", label, state)
        return state

    def log_browser_console(self, label: str) -> None:
        """Intenta registrar logs de consola del navegador."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        try:
            entries = self.driver.get_log("browser")
        except WebDriverException as exc:
            self.logger.warning("Browser console logs unavailable [%s]: %s", label, exc)
            return

        if not entries:
            self.logger.info("No browser console entries captured [%s].", label)
            return

        for entry in entries:
            level = entry.get("level", "UNKNOWN")
            message = entry.get("message", "")
            self.logger.info("Browser console [%s] %s: %s", label, level, message)

    def is_login_screen_visible(self) -> bool:
        """Ayuda a saber si seguimos en la pantalla de login."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")
        return len(self.driver.find_elements(By.CSS_SELECTOR, "#txtUsername-inputEl")) > 0

    def close(self) -> None:
        """Cierra recursos del navegador."""
        if self.driver:
            self.driver.quit()
            self.driver = None

    def _pause(self, seconds: float | None = None) -> None:
        delay = seconds if seconds is not None else self.settings.browser_slow_mo_ms / 1000
        if delay > 0:
            time.sleep(delay)

    def _get_ready_state(self) -> str:
        if not self.driver:
            raise RuntimeError("Browser session not started.")
        return str(self.driver.execute_script("return document.readyState"))

    def _scroll_into_view(self, element: Any) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            element,
        )

    def _extract_row_text(self, row: Any, selector: str) -> str:
        elements = row.find_elements(By.CSS_SELECTOR, selector)
        if not elements:
            return ""
        return elements[0].text.strip()

    def _collect_grid_records_from_dom(self) -> list[TicketRecord]:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        rows = self.driver.execute_script(
            r"""
            const grid = document.querySelector('div.x-grid-view');
            if (!grid) {
                return [];
            }

            const normalizeLabel = (value) =>
                (value || '')
                    .replace(/\s+/g, ' ')
                    .trim()
                    .toLowerCase();

            const headerCandidates = Array.from(
                document.querySelectorAll(
                    [
                        'div.x-column-header',
                        'td.x-column-header',
                        'th'
                    ].join(', ')
                )
            );

            let createdColumnId = '';
            let createdHeaderLabel = '';
            const exactCreatedHeaderMatchers = ['fecha', 'created'];
            const fuzzyCreatedHeaderMatchers = ['created date', 'creation date'];

            for (const header of headerCandidates) {
                const textNode = header.querySelector('.x-column-header-text') || header;
                const headerText = normalizeLabel(textNode.textContent);
                if (!headerText) {
                    continue;
                }
                const isExactMatch = exactCreatedHeaderMatchers.includes(headerText);
                const isFuzzyMatch = fuzzyCreatedHeaderMatchers.some((matcher) =>
                    headerText.includes(matcher)
                );
                if (!isExactMatch && !isFuzzyMatch) {
                    continue;
                }

                const candidateColumnId = header.id || '';

                if (candidateColumnId) {
                    createdColumnId = candidateColumnId;
                    createdHeaderLabel = headerText;
                    break;
                }
            }

            const rowSelectors = [
                'table.x-grid-item.list-view-row.list-grid-row',
                'table.x-grid-item'
            ];

            let rowElements = [];
            for (const selector of rowSelectors) {
                rowElements = Array.from(grid.querySelectorAll(selector));
                if (rowElements.length) {
                    break;
                }
            }

            const readText = (row, selector) => {
                const el = row.querySelector(selector);
                return el ? (el.textContent || '').trim() : '';
            };

            const readCreated = (row) => {
                if (!createdColumnId) {
                    return '';
                }

                const selectors = [
                    `td[data-columnid="${createdColumnId}"] .text-datetime-date`,
                    `td[data-columnid="${createdColumnId}"]`,
                ];

                for (const selector of selectors) {
                    const value = readText(row, selector);
                    if (value) {
                        return value;
                    }
                }
                return '';
            };

            return rowElements.map((row, index) => ({
                row_index: index,
                created_column_id: createdColumnId,
                created_header_label: createdHeaderLabel,
                ticket_id: readText(row, "td[data-columnid$='_c16_0'] .ticket-id"),
                created: readCreated(row),
                payment_method: readText(row, "td[data-columnid$='_c8_6'] .list-cell span"),
                subject: readText(row, "td[data-columnid$='_c27_8'] .text-default"),
                company: readText(row, "td[data-columnid$='_c9_5'] .list-cell span"),
                ticket_type: readText(row, "td[data-columnid$='_c39_7'] .text-caption"),
                hda_status: readText(row, "td[data-columnid$='_c36_3'] .label-text"),
                displayed: !!(row.offsetWidth || row.offsetHeight || row.getClientRects().length),
                text_preview: (row.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 300),
            }));
            """
        )

        ticket_records: list[TicketRecord] = []
        created_column_id = ""
        created_header_label = ""
        for row in rows:
            created_column_id = str(row.get("created_column_id", "")).strip() or created_column_id
            created_header_label = str(row.get("created_header_label", "")).strip() or created_header_label
            ticket_records.append(
                TicketRecord(
                    ticket_id=str(row.get("ticket_id", "")).strip(),
                    created=str(row.get("created", "")).strip(),
                    payment_method=str(row.get("payment_method", "")).strip(),
                    subject=str(row.get("subject", "")).strip(),
                    company=str(row.get("company", "")).strip(),
                    ticket_type=str(row.get("ticket_type", "")).strip(),
                    hda_status=str(row.get("hda_status", "")).strip(),
                )
            )
        self.logger.info(
            "Created date column discovery | header=%s | data_columnid=%s",
            created_header_label or "<not found>",
            created_column_id or "<not found>",
        )
        return ticket_records

    def _get_grid_scroll_state(self) -> dict[str, Any]:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        return dict(
            self.driver.execute_script(
                """
                const grid = document.querySelector('div.x-grid-view');
                if (!grid) {
                    return { grid_found: false };
                }

                const isScrollable = (el) => {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const overflowY = `${style.overflowY} ${style.overflow}`;
                    return /(auto|scroll)/.test(overflowY) && el.scrollHeight > el.clientHeight + 5;
                };

                let candidate = grid;
                while (candidate) {
                    if (isScrollable(candidate)) {
                        return {
                            grid_found: true,
                            scroller_found: true,
                            tag: candidate.tagName,
                            class_name: candidate.className,
                            scroll_top: candidate.scrollTop,
                            scroll_height: candidate.scrollHeight,
                            client_height: candidate.clientHeight,
                        };
                    }
                    candidate = candidate.parentElement;
                }

                return {
                    grid_found: true,
                    scroller_found: false,
                    grid_class: grid.className,
                };
                """
            )
        )

    def _scroll_grid_container(self) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        result = self.driver.execute_script(
            """
            const grid = document.querySelector('div.x-grid-view');
            if (!grid) {
                return { moved: false, reason: 'grid_not_found' };
            }

            const isScrollable = (el) => {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const overflowY = `${style.overflowY} ${style.overflow}`;
                return /(auto|scroll)/.test(overflowY) && el.scrollHeight > el.clientHeight + 5;
            };

            let candidate = grid;
            while (candidate && !isScrollable(candidate)) {
                candidate = candidate.parentElement;
            }

            if (!candidate) {
                return { moved: false, reason: 'scroll_container_not_found' };
            }

            const before = candidate.scrollTop;
            const step = Math.max(Math.floor(candidate.clientHeight * 0.8), 120);
            const maxTop = Math.max(candidate.scrollHeight - candidate.clientHeight, 0);
            const target = Math.min(before + step, maxTop);
            candidate.scrollTop = target;

            return {
                moved: candidate.scrollTop > before,
                reason: candidate.scrollTop > before ? 'scrolled' : 'end_reached',
                before,
                after: candidate.scrollTop,
                client_height: candidate.clientHeight,
                scroll_height: candidate.scrollHeight,
            };
            """
        )
        self.logger.info("Grid scroll result: %s", result)
        return bool(result.get("moved"))

    def _go_to_next_grid_page(self, current_page_ticket_ids: list[str], page_index: int) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        next_page_xpath = (
            "//div[contains(@class,'x-toolbar') or contains(@class,'x-pagingtoolbar')]"
            "//a[.//span[contains(@class, 'x-tbar-page-next')]]"
        )
        next_buttons = self.driver.find_elements(By.XPATH, next_page_xpath)
        if not next_buttons:
            self.logger.info(
                "Grid pagination | page=%s | next_button_found=False",
                page_index,
            )
            return False

        next_button = next_buttons[0]
        classes = (next_button.get_attribute("class") or "").strip()
        aria_disabled = (next_button.get_attribute("aria-disabled") or "").strip().lower()
        is_disabled = (
            "x-item-disabled" in classes
            or "disabled" in classes.lower()
            or aria_disabled == "true"
        )
        self.logger.info(
            "Grid pagination | page=%s | next_button_found=True | disabled=%s | classes=%s | aria_disabled=%s",
            page_index,
            is_disabled,
            classes or "<none>",
            aria_disabled or "<none>",
        )
        if is_disabled:
            return False

        before_signature = tuple(current_page_ticket_ids[:5])
        self._scroll_into_view(next_button)
        self.driver.execute_script("arguments[0].click();", next_button)
        self.logger.info("Grid pagination | clicked next page button from page %s", page_index)

        try:
            WebDriverWait(self.driver, 15).until(
                lambda _: self._grid_page_changed(before_signature)
            )
            self._pause(1.0)
            return True
        except TimeoutException:
            self.logger.warning(
                "Grid pagination | next page click from page %s did not produce a detectable page change.",
                page_index,
            )
            return False

    def _grid_page_changed(self, previous_signature: tuple[str, ...]) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        current_rows = self._collect_grid_records_from_dom()
        current_signature = tuple(
            row.ticket_id for row in current_rows if row.ticket_id
        )[:5]
        return bool(current_signature) and current_signature != previous_signature

    def _log_grid_debug(self) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        grid_views = self.driver.find_elements(By.CSS_SELECTOR, "div.x-grid-view")
        self.logger.info("Grid debug | grid_views_found=%s", len(grid_views))

        for index, grid in enumerate(grid_views[:3], start=1):
            try:
                classes = grid.get_attribute("class")
                text_preview = " ".join(grid.text.split())[:400]
                self.logger.info(
                    "Grid debug view %s | classes=%s | text_preview=%s",
                    index,
                    classes,
                    text_preview or "<empty>",
                )
            except Exception as exc:
                self.logger.warning("Grid debug view %s could not be inspected: %s", index, exc)

        row_candidates = self.driver.find_elements(By.CSS_SELECTOR, "table.x-grid-item")
        self.logger.info("Grid debug | generic_row_candidates=%s", len(row_candidates))
        for index, row in enumerate(row_candidates[:5], start=1):
            try:
                text_preview = " ".join(row.text.split())[:300]
                classes = row.get_attribute("class")
                displayed = row.is_displayed()
                self.logger.info(
                    "Grid debug row %s | displayed=%s | classes=%s | text_preview=%s",
                    index,
                    displayed,
                    classes,
                    text_preview or "<empty>",
                )
            except Exception as exc:
                self.logger.warning("Grid debug row %s could not be inspected: %s", index, exc)

    def _ensure_payments_tab_active(self) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if self._is_payments_tab_active():
            return

        payments_tabs = self.driver.find_elements(
            By.XPATH,
            "//a[contains(@class,'x-tab') and .//span[normalize-space()='Payments']]",
        )
        if not payments_tabs:
            return

        payments_tab = payments_tabs[0]
        self.logger.info("Activating Payments tab before interacting with the grid.")
        self._scroll_into_view(payments_tab)
        self.driver.execute_script("arguments[0].click();", payments_tab)
        self._pause(1.5)

    def _is_payments_tab_active(self) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        active_payments_tabs = self.driver.find_elements(
            By.XPATH,
            "//a[contains(@class,'x-tab-active') and .//span[normalize-space()='Payments']]",
        )
        return len(active_payments_tabs) > 0

    def _is_ticket_detail_open(self, ticket_id: str | None = None) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        current_url = self.driver.current_url
        if "/WSCView/Detail/" in current_url:
            if not ticket_id:
                return True
            return ticket_id in current_url

        if not ticket_id:
            return False

        active_ticket_tabs = self.driver.find_elements(
            By.XPATH,
            f"//a[contains(@class,'x-tab-active') and contains(@aria-label,'{ticket_id}')]",
        )
        return len(active_ticket_tabs) > 0

    def _wait_for_ticket_detail(self, ticket_id: str, timeout: int = 10) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        try:
            WebDriverWait(self.driver, timeout).until(lambda _: self._is_ticket_detail_open(ticket_id))
            return True
        except TimeoutException:
            return False

    def _open_ticket_from_context_menu(self, row: Any, ticket_id: str) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        ctx_menu_button = row.find_element(
            By.CSS_SELECTOR,
            "div.gridCtxMenu.js-list-ctx-menu",
        )
        self._scroll_into_view(ctx_menu_button)
        self.driver.execute_script("arguments[0].click();", ctx_menu_button)
        self.logger.info("Context menu button clicked for ticket: %s", ticket_id)
        self._pause(1.0)

        open_option = WebDriverWait(self.driver, 20).until(
            ec.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[contains(@class,'x-menu') and (@aria-hidden='false' or contains(@style,'visibility: visible'))]//span[contains(@class,'x-menu-item-text') and normalize-space()='Abrir']",
                )
            )
        )
        self.driver.execute_script("arguments[0].click();", open_option)
        self.logger.info("Open menu option clicked for ticket: %s", ticket_id)

    def _open_ticket_with_enter(self, row: Any) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        ActionChains(self.driver).move_to_element(row).click(row).send_keys("\ue007").perform()

    def fetch_payment_tickets(self) -> list:
        raise NotImplementedError

    def download_payment_request_pdf(self, ticket_id: str) -> str:
        raise NotImplementedError

