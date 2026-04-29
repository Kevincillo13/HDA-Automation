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
        # Cambiado a SEVERE para silenciar la basura de ExtJS en la terminal
        options.set_capability("goog:loggingPrefs", {"browser": "SEVERE", "performance": "INFO"})
        options.add_argument("--log-level=3")
        
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
        """Lee las filas del grid de Payments asegurando visibilidad."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        self._ensure_payments_tab_active()
        self._pause(1.5)

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
        
        for index, record in enumerate(ticket_records, start=1):
            self.logger.info(
                "Grid row %s | ticket_id=%s | status=%s | created=%s | payment_method=%s | company=%s | type=%s | subject=%s",
                index,
                record.ticket_id,
                record.status,
                record.created,
                record.payment_method,
                record.company,
                record.ticket_type,
                record.subject,
            )

    # Filtro blindado contra espacios invisibles y mayúsculas
            otc_records = [
                ticket for ticket in ticket_records 
                if "onetime check" in ticket.payment_method.replace('\xa0', ' ').strip().lower()
                and ticket.status.replace('\xa0', ' ').strip().lower() in ["open", "assigned", "suspended"]
            ]
            self.logger.info("OneTime Check candidates visible: %s", len(otc_records))

            return otc_records

    def open_ticket_by_id(self, ticket_id: str) -> None:
        """Abre un ticket desde el grid con varias estrategias."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if self._is_ticket_detail_open(ticket_id):
            self.logger.info("Ticket already open in detail view: %s", ticket_id)
            return

        self._ensure_payments_tab_active()
        wait = WebDriverWait(self.driver, 30)
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
        """Cierra la pestaña del ticket activo."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        self.logger.info("Closing active ticket tab.")
        try:
            active_tab_selector = (
                By.XPATH,
                "//a[contains(@class, 'x-tab-active') and not(.//span[normalize-space()='Payments'])]",
            )
            active_tab = WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located(active_tab_selector)
            )

            close_button = active_tab.find_element(By.CSS_SELECTOR, "span.x-tab-close-btn")
            self._scroll_into_view(close_button)
            self.driver.execute_script("arguments[0].click();", close_button)
            self.logger.info("Active ticket tab close button clicked.")

            WebDriverWait(self.driver, 15).until(
                lambda _: self._is_payments_tab_active()
            )
            self.logger.info("Successfully returned to Payments tab.")
            self._pause(1.0)
        except Exception as exc:
            self.logger.warning("Could not seamlessly close tab: %s", exc)
            self._ensure_payments_tab_active()

    def suspend_ticket(self, ticket_id: str, reasons: list[str]) -> None:
        """Versión simplificada para pruebas: Solo loguea la intención."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")
            
        self.logger.info(f"TEST MODE: Would have suspended ticket {ticket_id} for reasons: {reasons}")
        self._pause(1.0)

    def _collect_grid_records_from_dom(self) -> list[TicketRecord]:
        """Extrae los datos basándose en visibilidad y previene sobrescritura de IDs."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        rows = self.driver.execute_script(
            r"""
            const allViews = Array.from(document.querySelectorAll('div.x-grid-view'));
            const activeView = allViews.find(v => v.offsetParent !== null);
            if (!activeView) return [];

            const panel = activeView.closest('.x-panel');
            if (!panel) return [];

            const headers = Array.from(panel.querySelectorAll('.x-column-header'));
            const colMap = {};
            
            const targets = {
                ticket_id: ['id', 'ticket', 'nº', 'numero'],
                created: ['date', 'fecha', 'created', 'creación'],
                payment_method: ['payment method', 'metodo de pago', 'método de pago', 'payment'],
                subject: ['subject', 'asunto', 'objeto'],
                company: ['company', 'empresa', 'sociedad', 'soc.'],
                ticket_type: ['type', 'tipo'],
                status: ['status', 'estado']
            };

            // EL CANDADO BLINDADO
            headers.forEach(h => {
                const textEl = h.querySelector('.x-column-header-text');
                if (!textEl) return;
                
                const text = textEl.textContent.toLowerCase().trim();
                if (!text) return;

                let colId = h.getAttribute('data-columnid') || h.id;

                for (const [key, aliases] of Object.entries(targets)) {
                    if (!colMap[key]) {
                        if (aliases.includes(text)) {
                            colMap[key] = colId; 
                        } else if (aliases.some(a => text.includes(a) && a.length > 3)) {
                            colMap[key] = colId; 
                        }
                    }
                }
            });

            const rowElements = Array.from(activeView.querySelectorAll('table.x-grid-item'));
            
            return rowElements.map(row => {
                const getCellText = (key, subSelector) => {
                    const cid = colMap[key];
                    if (!cid) return '';
                    const cell = row.querySelector(`td[data-columnid="${cid}"]`);
                    if (!cell) return '';
                    if (subSelector) {
                        const subEl = cell.querySelector(subSelector);
                        return subEl ? subEl.textContent.trim() : cell.textContent.trim();
                    }
                    return cell.textContent.trim();
                };

                return {
                    ticket_id: getCellText('ticket_id', '.ticket-id'),
                    created: getCellText('created', null),
                    payment_method: getCellText('payment_method', 'span'),
                    subject: getCellText('subject', '.text-default'),
                    company: getCellText('company', 'span'),
                    ticket_type: getCellText('ticket_type', '.text-caption'),
                    status: getCellText('status', '.label-text')
                };
            });
            """
        )

        ticket_records: list[TicketRecord] = []
        for row in (rows or []):
            ticket_records.append(
                TicketRecord(
                    ticket_id=str(row.get("ticket_id", "")).strip(),
                    created=str(row.get("created", "")).strip(),
                    status=str(row.get("status", "")).strip(),
                    payment_method=str(row.get("payment_method", "")).strip(),
                    subject=str(row.get("subject", "")).strip(),
                    company=str(row.get("company", "")).strip(),
                    ticket_type=str(row.get("ticket_type", "")).strip(),
                )
            )
        return ticket_records

    def _get_grid_scroll_state(self) -> dict[str, Any]:
        """Encuentra el estado del scroll usando la lógica de visibilidad nativa."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        return dict(
            self.driver.execute_script(
                """
                const allViews = Array.from(document.querySelectorAll('div.x-grid-view'));
                const grid = allViews.find(v => v.offsetParent !== null);
                
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
        """Hace scroll basándose en la visibilidad nativa del DOM."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        result = self.driver.execute_script(
            """
            const allViews = Array.from(document.querySelectorAll('div.x-grid-view'));
            const grid = allViews.find(v => v.offsetParent !== null);
            
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
            return False

        next_button = next_buttons[0]
        classes = (next_button.get_attribute("class") or "").strip()
        aria_disabled = (next_button.get_attribute("aria-disabled") or "").strip().lower()
        is_disabled = (
            "x-item-disabled" in classes
            or "disabled" in classes.lower()
            or aria_disabled == "true"
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
            return False

    def _grid_page_changed(self, previous_signature: tuple[str, ...]) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        current_rows = self._collect_grid_records_from_dom()
        current_signature = tuple(
            row.ticket_id for row in current_rows if row.ticket_id
        )[:5]
        return bool(current_signature) and current_signature != previous_signature

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

    def _open_ticket_with_enter(self, row: Any) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        ActionChains(self.driver).move_to_element(row).click(row).send_keys("\ue007").perform()

    def take_screenshot(self, name: str) -> Path:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if not self.settings.evidence_enabled:
            return Path(self.settings.evidence_dir) / f"{get_run_id()}_{name}.png"

        safe_name = f"{name}.png"
        path = Path(self.settings.evidence_dir) / f"{get_run_id()}_{safe_name}"
        self.driver.save_screenshot(str(path))
        return path

    def save_page_source(self, name: str) -> Path:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        if not self.settings.evidence_enabled:
            return Path(self.settings.evidence_dir) / f"{get_run_id()}_{name}.html"

        safe_name = f"{name}.html"
        path = Path(self.settings.evidence_dir) / f"{get_run_id()}_{safe_name}"
        path.write_text(self.driver.page_source, encoding="utf-8")
        return path

    def log_debug_state(self, label: str) -> dict[str, Any]:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        try:
            window_rect = self.driver.get_window_rect()
        except Exception as exc:
            window_rect = {"error": str(exc)}

        state = {
            "label": label,
            "url": self.driver.current_url,
            "title": self.driver.title,
            "active_payments_tab": self._is_payments_tab_active(),
        }
        self.logger.info("Browser state [%s]: %s", label, state)
        return state

    def log_browser_console(self, label: str) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        try:
            entries = self.driver.get_log("browser")
        except WebDriverException:
            return

        for entry in entries:
            level = entry.get("level", "UNKNOWN")
            message = entry.get("message", "")
            self.logger.info("Browser console [%s] %s: %s", label, level, message)

    def is_login_screen_visible(self) -> bool:
        if not self.driver:
            raise RuntimeError("Browser session not started.")
        return len(self.driver.find_elements(By.CSS_SELECTOR, "#txtUsername-inputEl")) > 0

    def close(self) -> None:
        if self.driver:
            self.driver.quit()
            self.driver = None

    def _pause(self, seconds: float | None = None) -> None:
        delay = seconds if seconds is not None else self.settings.browser_slow_mo_ms / 1000
        if delay > 0:
            time.sleep(delay)

    def _scroll_into_view(self, element: Any) -> None:
        if not self.driver:
            raise RuntimeError("Browser session not started.")
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
            element,
        )