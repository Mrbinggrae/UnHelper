from __future__ import annotations

import re
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Iterable

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement


LogCallback = Callable[[str], None]
Locator = tuple[str, str]


class WMSWeightError(RuntimeError):
    """Raised when a WMS product weight cannot be read safely."""


@dataclass(frozen=True, slots=True)
class ProductWeightLookup:
    sku_id: str
    product_name: str
    weight_grams: Decimal


@dataclass(frozen=True, slots=True)
class _SearchResult:
    sku_id: str
    product_name: str
    link: WebElement
    href: str

    @property
    def signature(self) -> tuple[str, str, str]:
        return self.sku_id, self.product_name, self.href


@dataclass(frozen=True, slots=True)
class _DetailContext:
    original_handle: str
    detail_handle: str
    opened_new_tab: bool
    original_url: str


class ProductWeightCrawler:
    """Read per-item weight from WMS product management in a dedicated Chrome."""

    WMS_URL = "https://wms.coupang.com/"
    LOGIN_NOTICE_INTERVAL_SECONDS = 30
    SKU_SEARCH_TIMEOUT_SECONDS = 30
    POLL_INTERVAL_SECONDS = 0.25

    FIRST_AUTH_LOCATOR: Locator = (By.ID, "xauth-wms-first-btn-span")
    USER_ID_LOCATOR: Locator = (By.ID, "username")
    PASSWORD_LOCATOR: Locator = (By.ID, "password")
    LOGIN_BUTTON_LOCATOR: Locator = (By.ID, "kc-login")

    INVENTORY_MENU_LOCATORS: tuple[Locator, ...] = (
        (
            By.XPATH,
            "//a[.//span[normalize-space()='재고관리'] or contains(normalize-space(.),'재고관리')]",
        ),
        (
            By.XPATH,
            "//a[contains(translate(@href,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'inventory')]",
        ),
        (By.XPATH, "/html/body/div[1]/div[2]/ul/li[6]/a"),
    )
    PRODUCT_MENU_LOCATORS: tuple[Locator, ...] = (
        (
            By.XPATH,
            "//a[.//span[normalize-space()='상품관리'] or contains(normalize-space(.),'상품관리')]",
        ),
        (
            By.XPATH,
            "//a[contains(translate(@href,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sku') "
            "and not(ancestor::table)]",
        ),
        (By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[2]/ul/li[1]/a"),
    )

    SKU_INPUT_LOCATOR: Locator = (
        By.CSS_SELECTOR,
        "input.form-control.input-external-id",
    )
    SKU_SEARCH_BUTTON_LOCATOR: Locator = (By.CSS_SELECTOR, "button.btn-search")
    RESULT_ROWS_LOCATOR: Locator = (By.CSS_SELECTOR, "#wms__container table tbody tr")
    HIDDEN_WEIGHT_LOCATOR: Locator = (By.CSS_SELECTOR, "input.hidden-weight")
    RESULT_EXTERNAL_ID_COLUMN_INDEX = 1
    RESULT_PRODUCT_NAME_COLUMN_INDEX = 10
    NO_RESULT_TEXT = "조회 결과가 없습니다"

    def __init__(
        self,
        chrome_driver_path: str | Path,
        user_id: str = "",
        user_password: str = "",
        *,
        stop_event: threading.Event | None = None,
        log: LogCallback | None = None,
        headless: bool = False,
        timeout: float = 15,
    ) -> None:
        self.chrome_driver_path = Path(chrome_driver_path)
        self.user_id = str(user_id or "")
        self.user_password = str(user_password or "")
        self.stop_event = stop_event or threading.Event()
        self.log = log or (lambda message: print(message))
        self.headless = headless
        self.timeout = max(float(timeout), 1.0)
        self.driver: webdriver.Chrome | None = None

    def start(self) -> None:
        """Start a new Chrome, authenticate, and enter WMS product management."""
        self._check_cancelled()
        if self.driver is not None:
            return

        self.driver = self._build_driver()
        try:
            self.log("WMS 페이지를 엽니다.")
            self._driver.get(self.WMS_URL)
            self._wait_document_ready(120)
            self._login_and_wait()
            self._navigate_to_product_management()
            self.log("WMS 상품관리 화면 준비가 완료되었습니다.")
        except WMSWeightError:
            raise
        except (NoSuchWindowException, TimeoutException, WebDriverException) as exc:
            # Keep the browser alive so the caller can save failure evidence,
            # then close it explicitly in its finally block.
            raise WMSWeightError(f"WMS 로그인 또는 상품관리 진입에 실패했습니다: {exc}") from exc

    def lookup(self, sku: object) -> ProductWeightLookup:
        """Look up one exact WMS SKU and return its positive weight in grams."""
        self._check_cancelled()
        if self.driver is None:
            raise WMSWeightError("WMS Chrome이 시작되지 않았습니다. start()를 먼저 호출해 주세요.")

        sku_id = self.normalize_sku(sku)
        if not sku_id:
            raise WMSWeightError("조회할 SKU ID가 비어 있습니다.")

        if not self._visible_elements(*self.SKU_INPUT_LOCATOR):
            self._navigate_to_product_management()

        detail: _DetailContext | None = None

        try:
            self.log(f"WMS 상품관리에서 SKU {sku_id}를 조회합니다.")
            result = self._search_exact_sku(sku_id)
            try:
                detail = self._open_product_detail(result.link)
            except StaleElementReferenceException:
                self.log("상품조회 결과가 갱신되어 SKU를 한 번 더 조회합니다.")
                result = self._search_exact_sku(sku_id)
                detail = self._open_product_detail(result.link)

            weight_grams = self._extract_weight_grams()
            self.log(
                f"SKU {sku_id} 무게 확인 완료: {self.format_decimal(weight_grams)}g / "
                f"{result.product_name}"
            )
            return ProductWeightLookup(
                sku_id=sku_id,
                product_name=result.product_name,
                weight_grams=weight_grams,
            )
        except WMSWeightError:
            raise
        except (NoSuchWindowException, TimeoutException, WebDriverException) as exc:
            raise WMSWeightError(f"SKU {sku_id} WMS 무게 조회에 실패했습니다: {exc}") from exc
        finally:
            if detail is not None:
                self._restore_after_detail(detail)

    def close(self) -> None:
        driver, self.driver = self.driver, None
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                pass

    def save_failure_evidence(
        self,
        output_dir: str | Path,
        exc: Exception,
    ) -> tuple[Path, ...]:
        """Save a screenshot, redacted HTML, and redacted failure summary."""
        root = Path(output_dir).expanduser()
        try:
            root.mkdir(parents=True, exist_ok=True)
        except OSError as save_error:
            self.log(f"WMS 오류 증거 폴더를 만들지 못했습니다: {save_error}")
            return ()

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        saved: list[Path] = []
        driver = self.driver

        if driver is not None and self._is_logged_in():
            screenshot_path = root / f"UnHelper_WMS_error_{stamp}.png"
            try:
                driver.save_screenshot(str(screenshot_path))
                saved.append(screenshot_path)
                self.log(f"WMS 오류 화면 저장: {screenshot_path}")
            except (OSError, WebDriverException):
                pass

        if driver is not None:
            html_path = root / f"UnHelper_WMS_error_{stamp}.html"
            try:
                html = self.redact_sensitive_html(driver.page_source)
                html_path.write_text(html, encoding="utf-8")
                saved.append(html_path)
                self.log(f"WMS 오류 HTML 저장: {html_path}")
            except (OSError, WebDriverException):
                pass

        summary_path = root / f"UnHelper_WMS_error_{stamp}.txt"
        summary = "\n".join(
            (
                f"오류: {exc}",
                f"주소: {self._safe_current_url()}",
                f"제목: {self._safe_title()}",
            )
        )
        try:
            summary_path.write_text(self.redact_sensitive_text(summary), encoding="utf-8")
            saved.append(summary_path)
        except OSError:
            pass
        return tuple(saved)

    @staticmethod
    def normalize_sku(value: object) -> str:
        text = re.sub(r"\s+", "", str(value or "")).replace(",", "")
        if re.fullmatch(r"\d+\.0+", text):
            return text.split(".", 1)[0]
        return text

    @staticmethod
    def normalize_product_name(value: object) -> str:
        # Slash is ordinary product-name text. Only real whitespace is collapsed.
        return re.sub(r"\s+", " ", str(value or "")).strip()

    @staticmethod
    def format_decimal(value: Decimal) -> str:
        return format(value.normalize(), "f")

    def redact_sensitive_text(self, value: object) -> str:
        text = str(value or "")
        for secret in (self.user_password, self.user_id):
            if secret:
                text = text.replace(secret, "***")
        text = re.sub(r"(?i)(authorization\s*:\s*bearer\s+)[^\s<>&]+", r"\1***", text)
        text = re.sub(
            r"(?i)((?:password|passwd|pwd|token|secret|authorization|cookie)[\"']?\s*[:=]\s*[\"']?)[^\"'\s<>&]+",
            r"\1***",
            text,
        )
        return text

    def redact_sensitive_html(self, html: object) -> str:
        text = self.redact_sensitive_text(html)

        def redact_input(match: re.Match[str]) -> str:
            tag = match.group(0)
            lowered = tag.lower()
            if not (
                re.search(r"\btype\s*=\s*[\"']?password", lowered)
                or re.search(
                    r"\b(?:id|name|autocomplete)\s*=\s*(?:"
                    r"[\"'][^\"']*(?:user|login|pass|pwd|secret|token)[^\"']*[\"']"
                    r"|[^\s>]*(?:user|login|pass|pwd|secret|token)[^\s>]*)",
                    lowered,
                )
            ):
                return tag
            tag = re.sub(
                r"(?i)(\bvalue\s*=\s*)([\"']).*?\2",
                lambda value_match: f"{value_match.group(1)}{value_match.group(2)}***{value_match.group(2)}",
                tag,
            )
            return re.sub(r"(?i)(\bvalue\s*=\s*)([^\"'\s>]+)", r"\1***", tag)

        return re.sub(r"(?is)<input\b[^>]*>", redact_input, text)

    def _build_driver(self) -> webdriver.Chrome:
        if not self.chrome_driver_path.is_file():
            raise WMSWeightError(f"ChromeDriver를 찾을 수 없습니다: {self.chrome_driver_path}")

        options = webdriver.ChromeOptions()
        options.page_load_strategy = "eager"
        options.add_argument("--window-size=1600,1000")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-notifications")
        if self.headless:
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")

        try:
            driver = webdriver.Chrome(
                service=Service(executable_path=str(self.chrome_driver_path)),
                options=options,
            )
        except WebDriverException as exc:
            raise WMSWeightError(
                "WMS용 Chrome을 시작하지 못했습니다. Chrome과 공용 ChromeDriver의 "
                f"메이저 버전을 확인해 주세요. ({exc})"
            ) from exc
        driver.set_page_load_timeout(120)
        return driver

    def _login_and_wait(self) -> None:
        if self._is_logged_in():
            self.log("WMS 로그인 상태를 확인했습니다.")
            return

        first_auth = self._wait_for_any(
            (self.FIRST_AUTH_LOCATOR,),
            timeout=5,
            require_clickable=True,
            raise_on_timeout=False,
        )
        if first_auth is not None:
            self._click_element(first_auth, "WMS 1차 인증")

        login_form = self._wait_for_login_form_or_session(30)
        if login_form and self.user_id and self.user_password:
            user_input = self._first_visible(*self.USER_ID_LOCATOR, timeout=5)
            password_input = self._first_visible(*self.PASSWORD_LOCATOR, timeout=5)
            self._set_input_value(user_input, self.user_id, "WMS ID")
            self._set_input_value(password_input, self.user_password, "WMS 비밀번호")
            login_button = self._first_clickable(*self.LOGIN_BUTTON_LOCATOR, timeout=5)
            self._click_element(login_button, "WMS 로그인")
            self.log("저장된 WMS 계정으로 로그인을 요청했습니다.")
        elif not self._is_logged_in():
            self.log("WMS 자동 로그인 정보를 사용할 수 없어 브라우저에서 로그인을 기다립니다.")

        self.log("추가 인증을 포함한 WMS 로그인 완료까지 제한 없이 기다립니다.")
        last_notice = time.monotonic()
        while not self._is_logged_in():
            self._check_cancelled()
            self._ensure_browser_open()
            now = time.monotonic()
            if now - last_notice >= self.LOGIN_NOTICE_INTERVAL_SECONDS:
                self.log("WMS 로그인을 계속 기다리는 중입니다. 브라우저에서 인증을 완료해 주세요.")
                last_notice = now
            self.stop_event.wait(1.0)
        self.log("WMS 로그인 완료를 확인했습니다.")

    def _wait_for_login_form_or_session(self, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self._check_cancelled()
            if self._is_logged_in():
                return False
            if self._visible_elements(*self.USER_ID_LOCATOR) and self._visible_elements(
                *self.PASSWORD_LOCATOR
            ):
                return True
            self.stop_event.wait(self.POLL_INTERVAL_SECONDS)
        return False

    def _is_logged_in(self) -> bool:
        if self._visible_elements(*self.SKU_INPUT_LOCATOR):
            return True
        return any(self._visible_elements(*locator) for locator in self.INVENTORY_MENU_LOCATORS)

    def _navigate_to_product_management(self) -> None:
        self._check_cancelled()
        if self._visible_elements(*self.SKU_INPUT_LOCATOR):
            return

        inventory = self._wait_for_any(
            self.INVENTORY_MENU_LOCATORS,
            timeout=30,
            require_clickable=True,
        )
        self._click_element(inventory, "재고관리 메뉴")
        product_menu = self._wait_for_any(
            self.PRODUCT_MENU_LOCATORS,
            timeout=30,
            require_clickable=True,
        )
        self._click_element(product_menu, "상품관리 메뉴")
        self._wait_document_ready(60)
        self._first_visible(*self.SKU_INPUT_LOCATOR, timeout=30)

    def _search_exact_sku(self, sku_id: str) -> _SearchResult:
        sku_input = self._first_visible(*self.SKU_INPUT_LOCATOR, timeout=self.timeout)
        self._set_input_value(sku_input, sku_id, "SKU ID")

        before_signature = self._table_signature()
        before_rows = self._result_rows()
        before_first_row = before_rows[0] if before_rows else None
        self._arm_result_monitor()

        button = self._first_clickable(*self.SKU_SEARCH_BUTTON_LOCATOR, timeout=self.timeout)
        self._click_element(button, "상품관리 조회")
        return self._wait_for_fresh_result(sku_id, before_signature, before_first_row)

    def _wait_for_fresh_result(
        self,
        sku_id: str,
        before_signature: tuple[str, ...],
        before_first_row: WebElement | None,
    ) -> _SearchResult:
        deadline = time.monotonic() + max(self.timeout, self.SKU_SEARCH_TIMEOUT_SECONDS)
        transition_seen = False
        last_candidate_signature: tuple[str, str, str] | None = None
        candidate_stable_count = 0
        no_result_stable_count = 0

        while time.monotonic() < deadline:
            self._check_cancelled()
            self._ensure_browser_open()
            current_signature = self._table_signature()
            transition_seen = transition_seen or bool(
                current_signature != before_signature
                or self._is_stale(before_first_row)
                or self._result_mutation_count() > 0
            )

            if transition_seen:
                candidate = self._find_exact_result(sku_id)
                if candidate is not None:
                    if candidate.signature == last_candidate_signature:
                        candidate_stable_count += 1
                    else:
                        last_candidate_signature = candidate.signature
                        candidate_stable_count = 1
                    if candidate_stable_count >= 2:
                        return candidate
                else:
                    last_candidate_signature = None
                    candidate_stable_count = 0

                if self._has_no_result_message():
                    no_result_stable_count += 1
                    if no_result_stable_count >= 2:
                        raise WMSWeightError(f"WMS 상품관리에서 SKU {sku_id}를 찾지 못했습니다.")
                else:
                    no_result_stable_count = 0

            self.stop_event.wait(self.POLL_INTERVAL_SECONDS)

        if not transition_seen:
            raise WMSWeightError(
                f"SKU {sku_id} 조회 후 새 검색 결과가 생성되었는지 확인하지 못했습니다. 다시 시도해 주세요."
            )
        raise WMSWeightError(f"WMS 상품관리에서 SKU {sku_id} 검색 결과가 안정화되지 않았습니다.")

    def _find_exact_result(self, sku_id: str) -> _SearchResult | None:
        for row in self._result_rows():
            result = self._parse_result_row(row, sku_id)
            if result is not None:
                return result
        return None

    @classmethod
    def _parse_result_row(cls, row: WebElement, sku_id: str) -> _SearchResult | None:
        try:
            columns = row.find_elements(By.TAG_NAME, "td")
            if len(columns) <= cls.RESULT_PRODUCT_NAME_COLUMN_INDEX:
                return None
            row_external_id = cls.normalize_sku(
                columns[cls.RESULT_EXTERNAL_ID_COLUMN_INDEX].text
            )
            if row_external_id != sku_id:
                return None
            product_column = columns[cls.RESULT_PRODUCT_NAME_COLUMN_INDEX]
            links = product_column.find_elements(By.TAG_NAME, "a")
            if not links:
                return None
            link = links[0]
            product_name = cls.normalize_product_name(link.text)
            if not product_name:
                product_name = cls.normalize_product_name(product_column.text)
            if not product_name:
                return None
            return _SearchResult(
                sku_id=row_external_id,
                product_name=product_name,
                link=link,
                href=str(link.get_attribute("href") or ""),
            )
        except StaleElementReferenceException:
            return None

    def _open_product_detail(self, link: WebElement) -> _DetailContext:
        self._check_cancelled()
        original_handle = self._driver.current_window_handle
        original_url = self._safe_current_url()
        before_handles = set(self._driver.window_handles)
        self._click_element(link, "상품 상세")

        deadline = time.monotonic() + self.timeout
        detail: _DetailContext | None = None
        try:
            while time.monotonic() < deadline:
                self._check_cancelled()
                self._ensure_browser_open()
                handles = self._driver.window_handles
                new_handles = [handle for handle in handles if handle not in before_handles]
                if new_handles:
                    detail_handle = new_handles[0]
                    detail = _DetailContext(original_handle, detail_handle, True, original_url)
                    self._driver.switch_to.window(detail_handle)
                    self._wait_document_ready(self.timeout)
                    return detail

                current_url = self._safe_current_url()
                if current_url != original_url or self._present_elements(*self.HIDDEN_WEIGHT_LOCATOR):
                    detail = _DetailContext(original_handle, original_handle, False, original_url)
                    self._wait_document_ready(self.timeout)
                    return detail
                self.stop_event.wait(self.POLL_INTERVAL_SECONDS)

            raise WMSWeightError("WMS 상품 상세 화면이 열리지 않았습니다.")
        except Exception:
            if detail is not None:
                self._restore_after_detail(detail)
            else:
                self._close_unexpected_detail_tabs(before_handles, original_handle)
            raise

    def _extract_weight_grams(self) -> Decimal:
        weight_input = self._first_present(*self.HIDDEN_WEIGHT_LOCATOR, timeout=self.timeout)
        raw_value = str(weight_input.get_attribute("value") or "").strip().replace(",", "")
        if not raw_value:
            raise WMSWeightError("WMS 상품 상세의 무게 값이 비어 있습니다.")
        try:
            weight = Decimal(raw_value)
        except (InvalidOperation, ValueError) as exc:
            raise WMSWeightError(f"WMS 상품 무게를 숫자로 변환하지 못했습니다: {raw_value}") from exc
        if not weight.is_finite() or weight <= 0:
            raise WMSWeightError(f"WMS 상품 무게가 0 이하이거나 올바르지 않습니다: {raw_value}")
        return weight

    def _restore_after_detail(self, detail: _DetailContext) -> None:
        driver = self.driver
        if driver is None:
            return
        try:
            handles = driver.window_handles
            if detail.opened_new_tab:
                if detail.detail_handle in handles:
                    driver.switch_to.window(detail.detail_handle)
                    driver.close()
                handles = driver.window_handles
                if detail.original_handle in handles:
                    driver.switch_to.window(detail.original_handle)
                elif handles:
                    driver.switch_to.window(handles[0])
                return

            if detail.original_handle in handles:
                driver.switch_to.window(detail.original_handle)
            driver.back()
            try:
                self._wait_document_ready(self.timeout)
            except WMSWeightError:
                pass
        except (NoSuchWindowException, WebDriverException):
            pass

    def _close_unexpected_detail_tabs(
        self,
        before_handles: set[str],
        original_handle: str,
    ) -> None:
        driver = self.driver
        if driver is None:
            return
        try:
            for handle in list(driver.window_handles):
                if handle in before_handles:
                    continue
                driver.switch_to.window(handle)
                driver.close()
            handles = driver.window_handles
            if original_handle in handles:
                driver.switch_to.window(original_handle)
            elif handles:
                driver.switch_to.window(handles[0])
        except (NoSuchWindowException, WebDriverException):
            pass

    def _arm_result_monitor(self) -> None:
        try:
            self._driver.execute_script(
                """
                if (window.__unhelperWmsResultObserver) {
                  window.__unhelperWmsResultObserver.disconnect();
                }
                const container = document.querySelector('#wms__container');
                const table = container ? container.querySelector('table') : null;
                const root = table ? table.parentElement
                  : (container ? (container.children[1] || container) : document.body);
                window.__unhelperWmsResultState = { mutations: 0 };
                const observer = new MutationObserver((records) => {
                  window.__unhelperWmsResultState.mutations += records.length;
                });
                observer.observe(root, { childList: true, subtree: true, characterData: true });
                window.__unhelperWmsResultObserver = observer;
                """
            )
        except WebDriverException:
            pass

    def _result_mutation_count(self) -> int:
        try:
            value = self._driver.execute_script(
                "return (window.__unhelperWmsResultState || {}).mutations || 0;"
            )
            return int(value or 0)
        except (TypeError, ValueError, WebDriverException):
            return 0

    def _table_signature(self) -> tuple[str, ...]:
        signatures: list[str] = []
        for row in self._result_rows():
            try:
                signatures.append(self.normalize_product_name(row.text))
            except StaleElementReferenceException:
                continue
        return tuple(signatures)

    def _result_rows(self) -> list[WebElement]:
        try:
            return list(self._driver.find_elements(*self.RESULT_ROWS_LOCATOR))
        except (NoSuchWindowException, WebDriverException):
            return []

    def _has_no_result_message(self) -> bool:
        try:
            container_text = self._driver.find_element(By.ID, "wms__container").text
        except (NoSuchWindowException, WebDriverException):
            return False
        return self.NO_RESULT_TEXT in self.normalize_product_name(container_text)

    @staticmethod
    def _is_stale(element: WebElement | None) -> bool:
        if element is None:
            return False
        try:
            element.is_enabled()
            return False
        except StaleElementReferenceException:
            return True

    def _set_input_value(self, element: WebElement, value: str, label: str) -> None:
        target = str(value)
        for _ in range(2):
            self._check_cancelled()
            try:
                element.click()
                element.send_keys(Keys.CONTROL, "a")
                element.send_keys(Keys.DELETE)
                element.send_keys(target)
            except StaleElementReferenceException:
                raise
            except WebDriverException:
                pass
            if str(element.get_attribute("value") or "").strip() == target:
                return
            try:
                self._driver.execute_script(
                    """
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    """,
                    element,
                    target,
                )
            except WebDriverException:
                pass
            if str(element.get_attribute("value") or "").strip() == target:
                return
            self.stop_event.wait(0.2)
        raise WMSWeightError(f"{label} 입력값을 정확히 설정하지 못했습니다.")

    def _click_element(self, element: WebElement, label: str) -> None:
        self._check_cancelled()
        try:
            element.click()
        except StaleElementReferenceException:
            raise
        except WebDriverException:
            try:
                self._driver.execute_script("arguments[0].click();", element)
            except WebDriverException as exc:
                raise WMSWeightError(f"{label} 버튼을 클릭하지 못했습니다: {exc}") from exc

    def _wait_for_any(
        self,
        locators: Iterable[Locator],
        *,
        timeout: float,
        require_clickable: bool = False,
        raise_on_timeout: bool = True,
    ) -> WebElement | None:
        deadline = time.monotonic() + timeout
        locator_list = tuple(locators)
        while time.monotonic() < deadline:
            self._check_cancelled()
            for locator in locator_list:
                for element in self._present_elements(*locator):
                    try:
                        if not element.is_displayed():
                            continue
                        if require_clickable and not element.is_enabled():
                            continue
                        return element
                    except StaleElementReferenceException:
                        continue
            self.stop_event.wait(self.POLL_INTERVAL_SECONDS)
        if raise_on_timeout:
            raise WMSWeightError("WMS 화면에서 필요한 메뉴 또는 입력 요소를 찾지 못했습니다.")
        return None

    def _first_visible(self, by: str, selector: str, *, timeout: float) -> WebElement:
        element = self._wait_for_any(((by, selector),), timeout=timeout)
        if element is None:  # pragma: no cover - raise_on_timeout=True
            raise WMSWeightError(f"WMS 요소를 찾지 못했습니다: {selector}")
        return element

    def _first_clickable(self, by: str, selector: str, *, timeout: float) -> WebElement:
        element = self._wait_for_any(
            ((by, selector),),
            timeout=timeout,
            require_clickable=True,
        )
        if element is None:  # pragma: no cover - raise_on_timeout=True
            raise WMSWeightError(f"WMS 버튼을 찾지 못했습니다: {selector}")
        return element

    def _first_present(self, by: str, selector: str, *, timeout: float) -> WebElement:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self._check_cancelled()
            elements = self._present_elements(by, selector)
            if elements:
                return elements[0]
            self.stop_event.wait(self.POLL_INTERVAL_SECONDS)
        raise WMSWeightError(f"WMS 요소를 찾지 못했습니다: {selector}")

    def _present_elements(self, by: str, selector: str) -> list[WebElement]:
        try:
            return list(self._driver.find_elements(by, selector))
        except (NoSuchWindowException, WebDriverException):
            return []

    def _visible_elements(self, by: str, selector: str) -> list[WebElement]:
        visible: list[WebElement] = []
        for element in self._present_elements(by, selector):
            try:
                if element.is_displayed():
                    visible.append(element)
            except StaleElementReferenceException:
                continue
        return visible

    def _wait_document_ready(self, timeout: float) -> None:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self._check_cancelled()
            try:
                if self._driver.execute_script("return document.readyState") in {
                    "interactive",
                    "complete",
                }:
                    return
            except (NoSuchWindowException, WebDriverException):
                pass
            self.stop_event.wait(self.POLL_INTERVAL_SECONDS)
        raise WMSWeightError("WMS 페이지 로딩 대기 시간이 초과되었습니다.")

    def _ensure_browser_open(self) -> None:
        try:
            _ = self._driver.current_window_handle
        except (NoSuchWindowException, WebDriverException) as exc:
            raise WMSWeightError("WMS 로그인 대기 중 Chrome 창이 닫혔습니다.") from exc

    def _check_cancelled(self) -> None:
        if self.stop_event.is_set():
            raise WMSWeightError("사용자가 WMS 무게 조회를 중지했습니다.")

    def _safe_current_url(self) -> str:
        try:
            return str(self._driver.current_url)
        except (WMSWeightError, NoSuchWindowException, WebDriverException):
            return ""

    def _safe_title(self) -> str:
        try:
            return str(self._driver.title)
        except (WMSWeightError, NoSuchWindowException, WebDriverException):
            return ""

    @property
    def _driver(self) -> webdriver.Chrome:
        if self.driver is None:
            raise WMSWeightError("WMS Chrome이 시작되지 않았습니다.")
        return self.driver
