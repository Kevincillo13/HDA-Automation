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
        max_passes = 12

        for pass_index in range(1, max_passes + 1):
            snapshot = self._collect_grid_records_from_dom()
            self.logger.info(
                "Grid scan pass %s | dom_rows=%s | rows_with_ticket_id=%s",
                pass_index,
                len(snapshot),
                len([row for row in snapshot if row.ticket_id]),
            )

            for record in snapshot:
                if not record.ticket_id:
                    continue
                discovered_records.setdefault(record.ticket_id, record)

            moved = self._scroll_grid_container()
            self.logger.info(
                "Grid scan pass %s | unique_ticket_ids=%s | scroll_moved=%s",
                pass_index,
                len(discovered_records),
                moved,
            )
            if not moved:
                break
            self._pause(0.5)

        ticket_records = list(discovered_records.values())
        self.logger.info("Filtered grid rows with ticket IDs: %s", len(ticket_records))
        if not ticket_records:
            self._log_grid_debug()
        for index, record in enumerate(ticket_records, start=1):
            self.logger.info(
                "Grid row %s | ticket_id=%s | created=%s | payment_method=%s | company=%s | type=%s | subject=%s",
                index,
                record.ticket_id,
                record.created,
                record.payment_method,
                record.company,
                record.ticket_type,
                record.subject,
            )

        otc_records = [
            ticket for ticket in ticket_records if ticket.payment_method.strip() == "OneTime Check"
        ]
        self.logger.info("OneTime Check tickets visible: %s", len(otc_records))
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

    def reject_ticket(self, ticket_id: str, reason: str) -> None:
        """Rechaza un ticket en el portal. (Implementacion basica)."""
        if not self.driver:
            raise RuntimeError("Browser session not started.")

        self.logger.warning(
            "REJECTING TICKET | ticket_id=%s | reason=%s", ticket_id, reason
        )
        # TODO: Implementar la logica real para rechazar en la UI.
        # Por ahora, solo logeamos y cerramos la pestaña.
        self.logger.info("Ticket rejection logic not implemented. Skipping UI interaction.")
