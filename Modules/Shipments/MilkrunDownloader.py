from __future__ import annotations

import errno
import os
import re
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable

from selenium import webdriver
from selenium.common.exceptions import (
    JavascriptException,
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


LogCallback = Callable[[str], None]


class AutomationCancelled(RuntimeError):
    """Raised when the user stops an active browser workflow."""


@dataclass(frozen=True)
class MilkrunDownloadRequest:
    download_dir: Path
    center_name: str = "안산2"
    today: date | None = None


@dataclass(frozen=True)
class MilkrunDownloadResult:
    file_path: Path
    start_date: date
    end_date: date
    reason: str


@dataclass(frozen=True)
class HistoryEntry:
    row: WebElement | None
    download_type: str
    status: str
    reason: str
    requested_at: datetime | None
    index: int
    download_href: str = ""


class MilkrunDownloader:
    DASHBOARD_URL = (
        "https://shipments.coupang.net/dashboard?menuUrl=%2F&showAllMenus=true&locale=ko"
        "&assetBasePath=%2Fresources%2F20260722190911%2F&projectName=InboundShipment"
    )
    SCHEDULE_HREF = "/app"
    MILKRUN_LIST_HREF = "/app/inbound-booking/milkrun/list"
    BOOKING_LIST_HREF = MILKRUN_LIST_HREF
    BOOKING_LIST_LABEL = "밀크런 입고예약 목록"
    HISTORY_HREF = "/app/csv-download/history"
    DOWNLOAD_HREF_FRAGMENT = "/ibs/csv-donwload?"
    READY_STATUS = "다운로드 준비완료"
    DOWNLOAD_TYPE = "밀크런 입고예약 목록"
    LOGIN_LOG_INTERVAL_SECONDS = 30
    HISTORY_READY_TIMEOUT_SECONDS = 15 * 60
    DOWNLOAD_TIMEOUT_SECONDS = 5 * 60
    DOWNLOAD_EXTENSIONS = frozenset({".csv", ".txt", ".tsv", ".xlsx", ".xlsm", ".xlsb", ".xls"})
    STAGING_PREFIX = ".unhelper-download-"

    def __init__(
        self,
        chrome_driver_path: str | Path,
        log: LogCallback | None = None,
        stop_event: threading.Event | None = None,
        headless: bool = False,
    ):
        self.chrome_driver_path = Path(chrome_driver_path)
        self.log = log or (lambda message: print(message))
        self.stop_event = stop_event or threading.Event()
        self.headless = headless
        self.driver: webdriver.Chrome | None = None

    def run(
        self,
        request: MilkrunDownloadRequest,
        *,
        keep_browser_open: bool = False,
    ) -> MilkrunDownloadResult:
        target_date = request.today or date.today()
        start_date, end_date = self._resolve_date_range(target_date)
        reason = self.format_reason(start_date, end_date)
        download_dir = Path(request.download_dir).expanduser().resolve()
        download_dir.mkdir(parents=True, exist_ok=True)
        staging_dir = Path(
            tempfile.mkdtemp(prefix=self.STAGING_PREFIX, dir=str(download_dir))
        ).resolve()
        succeeded = False

        try:
            self.driver = self._build_driver(staging_dir)
            self._open_and_wait_for_login()
            self._open_booking_list()
            self._set_date_range(start_date, end_date)
            self._select_center(request.center_name)
            before_query = self._result_table_signature()
            self._click_button_text("조회", "조회 버튼")
            self.log("조회 결과가 표시될 때까지 기다립니다.")
            self._wait_for_query_complete(before_query)
            self._wait_for_text_download_button()
            self._click_text_download()
            known_history_hrefs = self._snapshot_history_download_hrefs()

            request_started_at = datetime.now()
            self._submit_download_reason(reason)
            self._close_request_confirmation()
            self._open_download_history()
            before_downloads = self._download_snapshot(staging_dir)
            self._download_latest_history_file(reason, request_started_at, known_history_hrefs)
            staged_file = self._wait_for_download(staging_dir, before_downloads)
            file_path = self._move_staged_download(staged_file, download_dir)
            self.log(f"다운로드 완료: {file_path}")
            result = MilkrunDownloadResult(
                file_path=file_path,
                start_date=start_date,
                end_date=end_date,
                reason=reason,
            )
            succeeded = True
            return result
        except AutomationCancelled:
            raise
        except Exception as exc:
            self._save_failure_snapshot(download_dir, exc)
            raise
        finally:
            if not keep_browser_open or not succeeded:
                self.close()
            self._cleanup_staging_dir(staging_dir, download_dir)

    def close(self) -> None:
        driver, self.driver = self.driver, None
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                pass

    def cancel(self) -> None:
        self.stop_event.set()

    def save_failure_snapshot(self, output_dir: str | Path, exc: Exception) -> None:
        """Persist browser evidence for failures in a follow-up authenticated step."""
        self._save_failure_snapshot(Path(output_dir).expanduser(), exc)

    @staticmethod
    def format_reason(start_date: date, end_date: date) -> str:
        return f"{start_date:%m.%d}-{end_date:%m.%d}"

    @classmethod
    def _resolve_date_range(cls, target_date: date) -> tuple[date, date]:
        return target_date - timedelta(days=1), target_date

    @staticmethod
    def parse_material_date_text(value: str) -> tuple[int, int, int] | None:
        numbers = [int(token) for token in re.findall(r"\d+", value or "")]
        if len(numbers) < 3:
            return None
        if numbers[0] >= 1000:
            return numbers[0], numbers[1], numbers[2]
        if numbers[2] >= 1000:
            return numbers[2], numbers[0], numbers[1]
        return None

    @classmethod
    def choose_latest_history_entry(
        cls,
        entries: Iterable[HistoryEntry],
        reason: str,
        requested_after: datetime,
        excluded_hrefs: set[str] | None = None,
    ) -> HistoryEntry | None:
        # The site exposes request time only to the minute, and the server clock
        # can differ slightly from the PC. Unique download URLs are the primary
        # discriminator; this time window is only a secondary guard.
        threshold = requested_after.replace(second=0, microsecond=0) - timedelta(minutes=5)
        excluded = excluded_hrefs or set()
        matching = [
            entry
            for entry in entries
            if entry.download_type.strip() == cls.DOWNLOAD_TYPE
            and entry.reason.strip() == reason
            and entry.requested_at is not None
            and entry.requested_at >= threshold
            and (not entry.download_href or entry.download_href not in excluded)
        ]
        if not matching:
            return None
        return max(matching, key=lambda entry: (entry.requested_at, entry.index))

    def _build_driver(self, download_dir: Path) -> webdriver.Chrome:
        if not self.chrome_driver_path.is_file():
            raise FileNotFoundError(f"ChromeDriver를 찾을 수 없습니다: {self.chrome_driver_path}")

        options = webdriver.ChromeOptions()
        options.page_load_strategy = "eager"
        options.add_argument("--window-size=1600,1000")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-notifications")
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(download_dir),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            },
        )
        if self.headless:
            options.add_argument("--headless=new")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")

        service = Service(executable_path=str(self.chrome_driver_path))
        try:
            driver = webdriver.Chrome(service=service, options=options)
        except WebDriverException as exc:
            raise RuntimeError(
                "Chrome을 시작하지 못했습니다. Chrome과 공용 ChromeDriver의 메이저 버전을 확인해 주세요. "
                f"({exc})"
            ) from exc
        driver.set_page_load_timeout(120)
        return driver

    def _open_and_wait_for_login(self) -> None:
        self.log("Coupang Shipments 페이지를 엽니다.")
        self._driver.get(self.DASHBOARD_URL)
        self._wait_document_ready(120)
        self.log("브라우저에서 직접 로그인해 주세요. 로그인 완료까지 제한 없이 기다립니다.")

        last_notice = time.monotonic()
        while True:
            self._check_cancelled()
            self._ensure_browser_open()
            if self._login_ready():
                self.log("로그인 완료를 감지했습니다.")
                return
            now = time.monotonic()
            if now - last_notice >= self.LOGIN_LOG_INTERVAL_SECONDS:
                self.log("로그인을 계속 기다리는 중입니다. 인증을 마치면 자동으로 이어집니다.")
                last_notice = now
            self.stop_event.wait(1.0)

    def _login_ready(self) -> bool:
        xpaths = [
            "//a[@href='/app' and .//span[normalize-space()='입고 스케줄']]",
            f"//a[@href='{self.BOOKING_LIST_HREF}' "
            f"and .//span[normalize-space()='{self.BOOKING_LIST_LABEL}']]",
        ]
        for xpath in xpaths:
            try:
                if any(element.is_displayed() for element in self._driver.find_elements(By.XPATH, xpath)):
                    return True
            except (StaleElementReferenceException, WebDriverException):
                continue
        return False

    def _open_booking_list(self) -> None:
        self.log("입고 스케줄 메뉴를 클릭합니다.")
        self._click_locator(
            By.XPATH,
            "//a[@href='/app' and .//span[normalize-space()='입고 스케줄']]",
            "입고 스케줄",
            timeout=60,
        )
        self._wait_document_ready(60)

        self.log(f"{self.BOOKING_LIST_LABEL}을 클릭합니다.")
        self._click_locator(
            By.XPATH,
            f"//a[@href='{self.BOOKING_LIST_HREF}' "
            f"and .//span[normalize-space()='{self.BOOKING_LIST_LABEL}']]",
            self.BOOKING_LIST_LABEL,
            timeout=60,
        )
        self._wait_document_ready(60)
        self._wait(
            60,
            lambda: bool(self._visible_elements(By.CSS_SELECTOR, "input[matstartdate], input[formcontrolname='startDate']")),
            "날짜 범위 입력",
        )

    def _set_date_range(self, start_date: date, end_date: date) -> None:
        self.log(f"조회 기간을 {start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}로 설정합니다.")
        start_input = self._first_visible(
            By.CSS_SELECTOR,
            "input[matstartdate], input[formcontrolname='startDate']",
            timeout=30,
        )
        end_input = self._first_visible(
            By.CSS_SELECTOR,
            "input[matenddate], input[formcontrolname='endDate']",
            timeout=30,
        )
        self._click_element(start_input, "시작 날짜")
        self._wait_for_calendar()
        self._select_calendar_date(start_date)

        if not self._visible_elements(By.CSS_SELECTOR, ".mat-datepicker-content, mat-calendar"):
            self._click_element(end_input, "종료 날짜")
            self._wait_for_calendar()
        self._select_calendar_date(end_date)

        self._wait(15, lambda: self._date_range_matches(start_date, end_date), "조회 날짜 범위 반영")

    def _wait_for_calendar(self) -> None:
        self._wait(
            15,
            lambda: bool(self._visible_elements(By.CSS_SELECTOR, ".mat-datepicker-content, mat-calendar")),
            "날짜 선택 달력",
        )

    def _select_calendar_date(self, target: date) -> None:
        for _ in range(36):
            self._check_cancelled()
            cells = self._visible_elements(By.CSS_SELECTOR, "button.mat-calendar-body-cell:not([disabled])")
            for cell in cells:
                if self._calendar_cell_matches(cell, target):
                    self._click_element(cell, f"{target:%Y-%m-%d}")
                    return

            delta = self._calendar_month_delta(target)
            if delta is None:
                raise RuntimeError(f"달력에서 {target:%Y-%m-%d} 날짜를 찾지 못했습니다.")
            selector = "button.mat-calendar-next-button" if delta > 0 else "button.mat-calendar-previous-button"
            before_period = self._calendar_period_text()
            self._click_locator(By.CSS_SELECTOR, selector, "달력 월 이동", timeout=10)
            self._wait(
                10,
                lambda: bool(self._calendar_period_text())
                and self._calendar_period_text() != before_period,
                "달력 월 전환",
            )
        raise RuntimeError(f"달력에서 {target:%Y-%m-%d}까지 이동하지 못했습니다.")

    def _calendar_cell_matches(self, cell: WebElement, target: date) -> bool:
        label = (cell.get_attribute("aria-label") or "").strip()
        parsed = self.parse_material_date_text(label)
        if parsed:
            return parsed == (target.year, target.month, target.day)

        english = {
            target.strftime("%B %d, %Y").replace(" 0", " "),
            target.strftime("%b %d, %Y").replace(" 0", " "),
        }
        return any(value.lower() in label.lower() for value in english)

    def _calendar_month_delta(self, target: date) -> int | None:
        text = self._calendar_period_text()
        if not text:
            return None
        parsed = self._parse_calendar_month(text)
        if parsed is None:
            return None
        year, month = parsed
        return (target.year - year) * 12 + target.month - month

    def _calendar_period_text(self) -> str:
        buttons = self._visible_elements(By.CSS_SELECTOR, "button.mat-calendar-period-button")
        return buttons[0].text.strip() if buttons else ""

    @staticmethod
    def _parse_calendar_month(text: str) -> tuple[int, int] | None:
        korean = re.search(r"(\d{4})\s*년\s*(\d{1,2})\s*월", text)
        if korean:
            return int(korean.group(1)), int(korean.group(2))
        numeric = re.search(r"(\d{4})\D+(\d{1,2})", text)
        if numeric:
            return int(numeric.group(1)), int(numeric.group(2))
        for fmt in ("%B %Y", "%b %Y"):
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.year, parsed.month
            except ValueError:
                continue
        return None

    def _date_input_matches(self, element: WebElement, target: date) -> bool:
        candidates = [element.get_attribute("value") or ""]
        try:
            wrapper = element.find_element(By.XPATH, "./ancestor::div[contains(@class,'mat-date-range-input-wrapper')][1]")
            candidates.extend(mirror.text for mirror in wrapper.find_elements(By.CSS_SELECTOR, ".mat-date-range-input-mirror"))
        except WebDriverException:
            pass
        return any(
            self.parse_material_date_text(value) == (target.year, target.month, target.day)
            for value in candidates
        )

    def _date_range_matches(self, start_date: date, end_date: date) -> bool:
        starts = self._visible_elements(
            By.CSS_SELECTOR,
            "input[matstartdate], input[formcontrolname='startDate']",
        )
        ends = self._visible_elements(
            By.CSS_SELECTOR,
            "input[matenddate], input[formcontrolname='endDate']",
        )
        return bool(
            starts
            and ends
            and self._date_input_matches(starts[0], start_date)
            and self._date_input_matches(ends[0], end_date)
        )

    def _select_center(self, center_name: str) -> None:
        self.log(f"물류 센터에서 '{center_name}'를 선택합니다.")
        selects = self._visible_elements(
            By.CSS_SELECTOR,
            "mat-select[formcontrolname='centerCode']",
        )
        if not selects:
            selects = self._visible_elements(By.CSS_SELECTOR, "mat-select[role='combobox'], mat-select")
        if not selects:
            raise RuntimeError("물류 센터 선택 목록을 찾지 못했습니다.")

        for select in selects:
            self._check_cancelled()
            try:
                if select.get_attribute("aria-disabled") == "true":
                    continue
                self._click_element(select, "물류 센터 목록")
                self._wait(
                    8,
                    lambda: bool(self._visible_elements(By.CSS_SELECTOR, "div[role='listbox'] mat-option")),
                    "물류 센터 옵션",
                )
                exact_options = self._driver.find_elements(
                    By.XPATH,
                    "//div[@role='listbox']//mat-option"
                    f"[.//span[contains(@class,'mdc-list-item__primary-text') and normalize-space()='{center_name}']]",
                )
                if exact_options:
                    self._click_element(exact_options[0], center_name)
                    self._wait(
                        10,
                        lambda: not self._visible_elements(By.CSS_SELECTOR, "div[role='listbox']"),
                        "물류 센터 선택 반영",
                    )
                    self._wait(
                        10,
                        lambda: center_name in self._normalize(select.text),
                        "선택된 물류 센터 표시",
                    )
                    return
                ActionChains(self._driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.2)
            except (TimeoutException, StaleElementReferenceException, WebDriverException):
                try:
                    ActionChains(self._driver).send_keys(Keys.ESCAPE).perform()
                except WebDriverException:
                    pass
        raise RuntimeError(f"물류 센터 목록에서 '{center_name}'를 찾지 못했습니다.")

    def _wait_for_text_download_button(self) -> None:
        self._wait(
            120,
            lambda: self._find_text_download_button() is not None,
            "텍스트 다운로드 버튼",
        )

    def _wait_for_query_complete(self, before_signature: str) -> None:
        started = time.monotonic()

        def query_started_or_settled() -> bool:
            if self._query_busy():
                return True
            signature = self._result_table_signature()
            if signature and signature != before_signature:
                return True
            # Some identical/no-result queries expose no durable loading marker.
            return time.monotonic() - started >= 4.0

        self._wait(30, query_started_or_settled, "조회 요청 시작")
        self._wait(120, lambda: not self._query_busy(), "조회 처리 완료")

        stable_signature = None
        stable_count = 0
        deadline = time.monotonic() + 8
        while time.monotonic() < deadline:
            self._check_cancelled()
            signature = self._result_table_signature()
            if signature == stable_signature:
                stable_count += 1
            else:
                stable_signature = signature
                stable_count = 1
            if stable_count >= 3:
                return
            self.stop_event.wait(0.35)

    def _query_busy(self) -> bool:
        return bool(
            self._visible_elements(
                By.CSS_SELECTOR,
                "mat-spinner, mat-progress-spinner, .mat-mdc-progress-spinner, [aria-busy='true']",
            )
        )

    def _result_table_signature(self) -> str:
        try:
            rows = self._driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            return "|".join(self._normalize(row.text) for row in rows[:8])
        except (StaleElementReferenceException, WebDriverException):
            return ""

    def _find_text_download_button(self) -> WebElement | None:
        xpath = (
            "//button[not(@disabled) and "
            "(contains(normalize-space(.), '텍스트 다운로드') or contains(@aria-label, '텍스트 다운로드') "
            "or contains(@title, '텍스트 다운로드'))]"
        )
        elements = self._visible_elements(By.XPATH, xpath)
        return elements[0] if elements else None

    def _click_text_download(self) -> None:
        self.log("텍스트 다운로드 버튼을 클릭합니다.")
        button = self._find_text_download_button()
        if button is None:
            raise RuntimeError("텍스트 다운로드 버튼을 찾지 못했습니다.")
        self._click_element(button, "텍스트 다운로드")

    def _submit_download_reason(self, reason: str) -> None:
        self.log(f"다운로드 사유를 입력합니다: {reason}")
        dialog = self._first_visible(By.CSS_SELECTOR, "mat-dialog-container", timeout=30)
        inputs = dialog.find_elements(By.CSS_SELECTOR, "input[matinput], input.mat-mdc-input-element")
        if not inputs:
            raise RuntimeError("다운로드 사유 입력란을 찾지 못했습니다.")
        reason_input = inputs[0]
        reason_input.clear()
        reason_input.send_keys(reason)
        self._wait(
            10,
            lambda: bool(
                self._visible_elements(
                    By.XPATH,
                    "//mat-dialog-container//button[not(@disabled) and contains(normalize-space(.), '다운로드 요청하기')]",
                )
            ),
            "다운로드 요청하기 버튼 활성화",
        )
        self._click_locator(
            By.XPATH,
            "//mat-dialog-container//button[not(@disabled) and contains(normalize-space(.), '다운로드 요청하기')]",
            "다운로드 요청하기",
            timeout=15,
        )

    def _close_request_confirmation(self) -> None:
        self._wait(
            60,
            lambda: bool(
                self._visible_elements(
                    By.XPATH,
                    "//mat-dialog-container//*[contains(normalize-space(.), '텍스트 다운로드 요청되었습니다.') ]",
                )
            ),
            "텍스트 다운로드 요청 완료 팝업",
        )
        self.log("요청 완료 팝업의 '닫다' 버튼을 클릭합니다.")
        self._click_locator(
            By.XPATH,
            "//mat-dialog-container//button[normalize-space(.)='닫다' or normalize-space(.)='닫기']",
            "요청 완료 팝업 닫기",
            timeout=30,
        )

    def _open_download_history(self) -> None:
        self.log("텍스트 다운로드 내역으로 이동합니다.")
        self._click_locator(
            By.CSS_SELECTOR,
            f"a[href='{self.HISTORY_HREF}']",
            "텍스트 다운로드 내역",
            timeout=60,
        )
        self._wait_document_ready(60)
        self._wait(
            60,
            lambda: bool(self._visible_elements(By.CSS_SELECTOR, "table tbody tr")),
            "텍스트 다운로드 내역 표",
        )

    def _snapshot_history_download_hrefs(self) -> set[str]:
        """Read existing download URLs in a temporary tab before submitting.

        The site records timestamps only to the minute and does not guarantee
        newest-first row order. Existing URL exclusion prevents a same-minute
        rerun from selecting an older ready file with the same reason.
        """
        self.log("방금 요청할 파일을 구분하기 위해 기존 다운로드 내역을 확인합니다.")
        driver = self._driver
        original_handle = driver.current_window_handle
        snapshot: set[str] = set()
        try:
            driver.switch_to.new_window("tab")
            driver.get(f"https://shipments.coupang.net{self.HISTORY_HREF}")
            self._wait_document_ready(60)
            self._wait(
                30,
                lambda: bool(self._visible_elements(By.CSS_SELECTOR, "table")),
                "기존 다운로드 내역 표",
            )
            for link in driver.find_elements(
                By.CSS_SELECTOR,
                f"table tbody a[href*='{self.DOWNLOAD_HREF_FRAGMENT}']",
            ):
                href = link.get_attribute("href") or ""
                if href:
                    snapshot.add(href)
            self.log(f"기존 다운로드 파일 {len(snapshot)}개를 확인했습니다.")
        except AutomationCancelled:
            raise
        except Exception as exc:
            # The main history pass still has type/reason/time guards, so a
            # snapshot failure is recoverable and should not abort the request.
            self.log(f"[경고] 기존 다운로드 내역 사전 확인을 건너뜁니다: {exc}")
        finally:
            try:
                if driver.current_window_handle != original_handle:
                    driver.close()
                driver.switch_to.window(original_handle)
            except WebDriverException:
                pass
        return snapshot

    def _download_latest_history_file(
        self,
        reason: str,
        requested_after: datetime,
        known_history_hrefs: set[str],
    ) -> None:
        deadline = time.monotonic() + self.HISTORY_READY_TIMEOUT_SECONDS
        last_log = 0.0
        last_refresh = time.monotonic()
        while time.monotonic() < deadline:
            self._check_cancelled()
            entries = self._read_history_entries()
            latest = self.choose_latest_history_entry(
                entries,
                reason,
                requested_after,
                known_history_hrefs,
            )
            if latest and latest.status.strip() == self.READY_STATUS and latest.row is not None:
                links = latest.row.find_elements(
                    By.CSS_SELECTOR,
                    f"a[href*='{self.DOWNLOAD_HREF_FRAGMENT}']",
                )
                if links:
                    self.log("방금 요청한 최신 파일의 다운로드하기 버튼을 클릭합니다.")
                    self._click_element(links[0], "다운로드하기")
                    return
            if latest and any(token in latest.status for token in ("실패", "오류", "취소")):
                raise RuntimeError(f"텍스트 다운로드 생성에 실패했습니다. 상태: {latest.status}")

            now = time.monotonic()
            if now - last_log >= 10:
                if latest:
                    self.log(f"다운로드 파일 준비를 기다리는 중입니다. 현재 상태: {latest.status or '확인 중'}")
                else:
                    self.log("방금 요청한 다운로드 내역이 나타나기를 기다리는 중입니다.")
                last_log = now
            self.stop_event.wait(2)
            self._check_cancelled()
            if time.monotonic() - last_refresh >= 15:
                self._driver.refresh()
                self._wait_document_ready(60)
                self._wait(
                    30,
                    lambda: bool(self._visible_elements(By.CSS_SELECTOR, "table tbody tr")),
                    "새로고침된 다운로드 내역 표",
                )
                last_refresh = time.monotonic()
        raise TimeoutException("15분 안에 방금 요청한 텍스트 파일이 준비되지 않았습니다.")

    def _read_history_entries(self) -> list[HistoryEntry]:
        entries: list[HistoryEntry] = []
        for index, row in enumerate(self._driver.find_elements(By.CSS_SELECTOR, "table tbody tr")):
            try:
                cells = row.find_elements(By.CSS_SELECTOR, "td")
                if len(cells) < 6:
                    continue
                texts = [self._normalize(cell.text) for cell in cells]
                requested_at = self._parse_requested_at(texts[3])
                links = cells[5].find_elements(By.CSS_SELECTOR, f"a[href*='{self.DOWNLOAD_HREF_FRAGMENT}']")
                href = links[0].get_attribute("href") if links else ""
                entries.append(
                    HistoryEntry(
                        row=row,
                        download_type=texts[0],
                        status=texts[1],
                        reason=texts[2],
                        requested_at=requested_at,
                        index=index,
                        download_href=href,
                    )
                )
            except StaleElementReferenceException:
                return []
        return entries

    @staticmethod
    def _parse_requested_at(value: str) -> datetime | None:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                continue
        return None

    def _wait_for_download(self, download_dir: Path, before: dict[Path, tuple[int, int]]) -> Path:
        self.log("Chrome 파일 다운로드 완료를 기다립니다.")
        deadline = time.monotonic() + self.DOWNLOAD_TIMEOUT_SECONDS
        stable: dict[Path, tuple[tuple[int, int], int]] = {}
        while time.monotonic() < deadline:
            self._check_cancelled()
            active_partials = []
            for partial in download_dir.glob("*.crdownload"):
                try:
                    stat = partial.stat()
                except FileNotFoundError:
                    # Chrome atomically renames .crdownload files at the end
                    # of a download. The name may disappear between glob()
                    # and stat(), which is a normal completion race.
                    continue
                signature = (stat.st_size, stat.st_mtime_ns)
                if partial not in before or before[partial] != signature:
                    active_partials.append(partial)
            candidates = []
            unsupported = []
            for path in download_dir.iterdir():
                if not path.is_file() or path.suffix.lower() == ".crdownload":
                    continue
                try:
                    stat = path.stat()
                except FileNotFoundError:
                    # A final name can also be replaced/renamed while the
                    # directory enumeration is in progress. Retry next poll.
                    continue
                signature = (stat.st_size, stat.st_mtime_ns)
                if path not in before or before[path] != signature:
                    collection = candidates if path.suffix.lower() in self.DOWNLOAD_EXTENSIONS else unsupported
                    collection.append((stat.st_mtime_ns, path, signature))

            for _mtime, path, signature in sorted(candidates, reverse=True):
                previous_signature, count = stable.get(path, ((-1, -1), 0))
                count = count + 1 if previous_signature == signature else 1
                stable[path] = (signature, count)
                if signature[0] > 0 and count >= 3 and not active_partials:
                    return path
            for _mtime, path, signature in sorted(unsupported, reverse=True):
                previous_signature, count = stable.get(path, ((-1, -1), 0))
                count = count + 1 if previous_signature == signature else 1
                stable[path] = (signature, count)
                if signature[0] > 0 and count >= 3 and not active_partials and not candidates:
                    raise RuntimeError(
                        f"다운로드된 파일 형식을 지원하지 않습니다: {path.name}"
                    )
            self.stop_event.wait(0.5)
        raise TimeoutException("5분 안에 Chrome 파일 다운로드가 완료되지 않았습니다.")

    @classmethod
    def _move_staged_download(cls, source: Path, download_dir: Path) -> Path:
        source = source.resolve()
        download_dir = download_dir.resolve()
        if (
            source.parent.parent != download_dir
            or not source.parent.name.startswith(cls.STAGING_PREFIX)
        ):
            raise RuntimeError("다운로드 임시 파일 경로가 올바르지 않습니다.")

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        for sequence in range(1000):
            if sequence == 0:
                destination = download_dir / source.name
            else:
                suffix = f"_{stamp}" if sequence == 1 else f"_{stamp}_{sequence}"
                destination = download_dir / f"{source.stem}{suffix}{source.suffix}"
            try:
                cls._rename_no_replace(source, destination)
                return destination
            except FileExistsError:
                continue
            except OSError as exc:
                if exc.errno == errno.EEXIST or getattr(exc, "winerror", None) in {80, 183}:
                    continue
                raise RuntimeError(f"다운로드 파일을 최종 폴더로 이동하지 못했습니다.\n{exc}") from exc
        raise RuntimeError("다운로드 파일의 중복 이름을 만들 수 없습니다.")

    @staticmethod
    def _rename_no_replace(source: Path, destination: Path) -> None:
        """Publish a staged file atomically without replacing an existing file."""
        if os.name == "nt":
            # Windows MoveFile semantics used by os.rename do not replace an
            # existing destination, so a collision is safely retryable.
            source.rename(destination)
            return

        # Keep tests and non-Windows development no-clobber as well. The
        # staging directory is a child of download_dir, so this is one volume.
        os.link(source, destination)
        try:
            source.unlink()
        except Exception:
            try:
                destination.unlink()
            except OSError:
                pass
            raise

    @classmethod
    def _cleanup_staging_dir(cls, staging_dir: Path, download_dir: Path) -> None:
        try:
            resolved_staging = staging_dir.resolve()
            resolved_download = download_dir.resolve()
            if (
                resolved_staging.parent != resolved_download
                or not resolved_staging.name.startswith(cls.STAGING_PREFIX)
            ):
                return
            for child in resolved_staging.iterdir():
                if child.is_file() or child.is_symlink():
                    child.unlink(missing_ok=True)
            resolved_staging.rmdir()
        except OSError:
            pass

    @staticmethod
    def _download_snapshot(download_dir: Path) -> dict[Path, tuple[int, int]]:
        snapshot: dict[Path, tuple[int, int]] = {}
        for path in download_dir.iterdir():
            if path.is_file():
                try:
                    stat = path.stat()
                except FileNotFoundError:
                    continue
                snapshot[path] = (stat.st_size, stat.st_mtime_ns)
        return snapshot

    def _click_button_text(self, text: str, label: str) -> None:
        self._click_locator(
            By.XPATH,
            f"//button[not(@disabled) and normalize-space(.)='{text}']",
            label,
            timeout=30,
        )

    def _click_locator(self, by: str, value: str, label: str, timeout: float) -> None:
        element = self._first_visible(by, value, timeout)
        self._click_element(element, label)

    def _click_element(self, element: WebElement, label: str) -> None:
        self._check_cancelled()
        try:
            self._driver.execute_script(
                "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                element,
            )
            element.click()
        except (WebDriverException, JavascriptException):
            try:
                self._driver.execute_script("arguments[0].click();", element)
            except WebDriverException as exc:
                raise RuntimeError(f"'{label}' 클릭에 실패했습니다: {exc}") from exc

    def _first_visible(self, by: str, value: str, timeout: float) -> WebElement:
        result = self._wait(
            timeout,
            lambda: next(iter(self._visible_elements(by, value)), None),
            value,
        )
        return result

    def _visible_elements(self, by: str, value: str) -> list[WebElement]:
        try:
            return [element for element in self._driver.find_elements(by, value) if element.is_displayed()]
        except (StaleElementReferenceException, WebDriverException):
            return []

    def _wait(self, timeout: float, condition, label: str):
        def checked(_driver):
            self._check_cancelled()
            return condition()

        try:
            return WebDriverWait(
                self._driver,
                timeout,
                poll_frequency=0.25,
                ignored_exceptions=(StaleElementReferenceException,),
            ).until(checked)
        except TimeoutException as exc:
            raise TimeoutException(f"'{label}' 대기 시간이 초과되었습니다.") from exc

    def _wait_document_ready(self, timeout: float) -> None:
        self._wait(
            timeout,
            lambda: self._driver.execute_script("return document.readyState") in {"interactive", "complete"},
            "페이지 로딩",
        )

    def _ensure_browser_open(self) -> None:
        try:
            _ = self._driver.current_window_handle
        except (NoSuchWindowException, WebDriverException) as exc:
            raise RuntimeError("로그인 대기 중 Chrome 창이 닫혔습니다.") from exc

    def _check_cancelled(self) -> None:
        if self.stop_event.is_set():
            raise AutomationCancelled("사용자가 작업을 중지했습니다.")

    @property
    def _driver(self) -> webdriver.Chrome:
        if self.driver is None:
            raise RuntimeError("ChromeDriver가 시작되지 않았습니다.")
        return self.driver

    @staticmethod
    def _normalize(value: str) -> str:
        return re.sub(r"\s+", " ", value or "").strip()

    def _save_failure_snapshot(self, download_dir: Path, exc: Exception) -> None:
        if self.driver is None:
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            screenshot = download_dir / f"UnHelper_error_{stamp}.png"
            self.driver.save_screenshot(str(screenshot))
            self.log(f"오류 화면 저장: {screenshot}")
        except WebDriverException:
            pass
        try:
            html = download_dir / f"UnHelper_error_{stamp}.html"
            html.write_text(self.driver.page_source, encoding="utf-8")
            self.log(f"오류 HTML 저장: {html}")
        except (OSError, WebDriverException):
            pass
        self.log(f"자동화 오류: {exc}")
