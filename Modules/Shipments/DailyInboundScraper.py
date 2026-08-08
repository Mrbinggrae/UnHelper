from __future__ import annotations

import sys
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING, Iterable

from selenium.common.exceptions import (
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement

from .DailyInbound import (
    MilkrunProductRow,
    normalize_dispatch_number,
    normalize_milkrun_card_number,
    parse_detail_table_cells,
)

if TYPE_CHECKING:
    from .MilkrunDownloader import MilkrunDownloader


class DailyInboundError(RuntimeError):
    """Raised when the daily inbound page cannot provide requested detail data."""

    def __init__(
        self,
        message: str,
        *,
        evidence_captured: bool = False,
        failure_url: str = "",
    ):
        super().__init__(message)
        self.evidence_captured = evidence_captured
        self.failure_url = failure_url


@dataclass(frozen=True)
class DailyInboundResult:
    products: tuple[MilkrunProductRow, ...]
    requested_dispatches: tuple[str, ...]
    matched_dispatches: tuple[str, ...]
    unmatched_dispatches: tuple[str, ...]


class DailyInboundScraper:
    DAILY_SCHEDULE_HREF = "/app/inbound-schedule"
    DETAIL_HREF_FRAGMENT = "/app/inbound-booking/milkrun/detail"

    def __init__(
        self,
        browser: MilkrunDownloader,
        *,
        evidence_dir: str | Path | None = None,
    ):
        self.browser = browser
        self.log = browser.log
        self.evidence_dir = Path(evidence_dir).expanduser() if evidence_dir else None

    def run(
        self,
        dispatch_numbers: Iterable[str],
        *,
        center_name: str,
        schedule_date: date,
    ) -> DailyInboundResult:
        requested = self._unique_dispatches(dispatch_numbers)
        if not requested:
            raise DailyInboundError(
                "다운로드 첫 시트의 A열에서 조회할 Milkrun 배차번호를 찾지 못했습니다. "
                "Excel 값 반영은 완료되었지만 일별 입고 상세 조회는 진행하지 않았습니다."
            )

        self._open_daily_schedule()
        self._select_center(center_name)
        self._set_schedule_date(schedule_date)
        self._query_schedule()

        products: list[MilkrunProductRow] = []
        matched: list[str] = []
        unmatched: list[str] = []
        seen_products: set[MilkrunProductRow] = set()

        for dispatch_number in requested:
            self.browser._check_cancelled()
            matching_count = len(self._matching_slots(dispatch_number))
            if matching_count == 0:
                unmatched.append(dispatch_number)
                self.log(f"일별 입고 카드에서 배차번호 {dispatch_number}를 찾지 못했습니다.")
                continue

            self.log(f"배차번호 {dispatch_number}의 상세 상품을 조회합니다.")
            matched.append(dispatch_number)
            for match_index in range(matching_count):
                self.browser._check_cancelled()
                cards = self._matching_slots(dispatch_number)
                if match_index >= len(cards):
                    raise DailyInboundError(
                        f"배차번호 {dispatch_number} 카드가 조회 중 변경되었습니다. 다시 실행해 주세요."
                    )
                try:
                    rows = self._open_detail_and_read(cards[match_index], dispatch_number)
                except Exception as exc:
                    if self.browser.stop_event.is_set():
                        raise
                    failure_url = str(getattr(exc, "failure_url", "")) or self._safe_current_url()
                    raise DailyInboundError(
                        f"배차번호 {dispatch_number} 상세 조회에 실패했습니다.\n"
                        f"실패 주소: {failure_url}\n{exc}",
                        evidence_captured=bool(
                            getattr(exc, "evidence_captured", False)
                        ),
                        failure_url=failure_url,
                    ) from exc
                for row in rows:
                    if row not in seen_products:
                        seen_products.add(row)
                        products.append(row)

        if not products:
            missing = ", ".join(unmatched or requested)
            raise DailyInboundError(
                "오늘 일별 입고 현황에서 표시할 상품 상세를 찾지 못했습니다.\n"
                f"미조회 배차번호: {missing}"
            )

        self.log(f"일별 입고 상세 {len(products)}개 상품을 수집했습니다.")
        return DailyInboundResult(
            products=tuple(products),
            requested_dispatches=requested,
            matched_dispatches=tuple(matched),
            unmatched_dispatches=tuple(unmatched),
        )

    @staticmethod
    def _unique_dispatches(values: Iterable[str]) -> tuple[str, ...]:
        seen: set[str] = set()
        result: list[str] = []
        for value in values:
            normalized = normalize_dispatch_number(value)
            if normalized and normalized not in seen:
                seen.add(normalized)
                result.append(normalized)
        return tuple(result)

    def _open_daily_schedule(self) -> None:
        self.log("일별 입고 현황 메뉴로 이동합니다.")
        self.browser._click_locator(
            By.XPATH,
            "//a[@href='/app/inbound-schedule' "
            "and .//span[normalize-space()='일별 입고 현황']]",
            "일별 입고 현황",
            timeout=60,
        )
        self.browser._wait_document_ready(60)
        self.browser._wait(
            60,
            lambda: self.DAILY_SCHEDULE_HREF in self._safe_current_url()
            and bool(
                self.browser._visible_elements(
                    By.CSS_SELECTOR,
                    "mat-select[formcontrolname='centerCode']",
                )
            )
            and bool(
                self.browser._visible_elements(
                    By.CSS_SELECTOR,
                    "input[formcontrolname='scheduleDate']",
                )
            ),
            "일별 입고 현황 화면",
        )

    def _select_center(self, center_name: str) -> None:
        self.log(f"일별 입고 센터를 '{center_name}'로 설정합니다.")
        selector = "mat-select[formcontrolname='centerCode']"
        select = self.browser._first_visible(By.CSS_SELECTOR, selector, timeout=30)
        if center_name in self.browser._normalize(select.text):
            return
        self.browser._click_element(select, "일별 입고 센터 목록")
        self.browser._wait(
            15,
            lambda: bool(self.browser._visible_elements(By.CSS_SELECTOR, "div[role='listbox'] mat-option")),
            "일별 입고 센터 옵션",
        )
        options = self.browser._driver.find_elements(
            By.XPATH,
            "//div[@role='listbox']//mat-option"
            f"[.//span[normalize-space()='{center_name}']]",
        )
        if not options:
            raise DailyInboundError(f"일별 입고 센터 목록에서 '{center_name}'를 찾지 못했습니다.")
        self.browser._click_element(options[0], center_name)
        self.browser._wait(
            15,
            lambda: any(
                center_name in self.browser._normalize(element.text)
                for element in self.browser._visible_elements(By.CSS_SELECTOR, selector)
            ),
            "일별 입고 센터 선택 반영",
        )

    def _set_schedule_date(self, target: date) -> None:
        self.log(f"일별 입고 조회 날짜를 {target:%Y-%m-%d}로 설정합니다.")
        selector = "input[formcontrolname='scheduleDate']"
        date_input = self.browser._first_visible(By.CSS_SELECTOR, selector, timeout=30)
        if self.browser._date_input_matches(date_input, target):
            return
        toggles = self.browser._visible_elements(
            By.CSS_SELECTOR,
            "mat-datepicker-toggle button[aria-label='Open calendar'], "
            "mat-datepicker-toggle button",
        )
        if toggles:
            self.browser._click_element(toggles[0], "일별 입고 조회 달력")
        else:
            self.browser._click_element(date_input, "일별 입고 조회 날짜")
        self.browser._wait_for_calendar()
        self.browser._select_calendar_date(target)
        self.browser._wait(
            15,
            lambda: any(
                self.browser._date_input_matches(element, target)
                for element in self.browser._visible_elements(By.CSS_SELECTOR, selector)
            ),
            "일별 입고 조회 날짜 반영",
        )

    def _query_schedule(self) -> None:
        before = self._slot_signature()
        self.log("일별 입고 스케줄을 조회합니다.")
        query_button = self._find_query_button()
        self._install_query_monitor()
        self._arm_query_monitor()
        transition_seen = False

        def query_transitioned() -> bool:
            nonlocal transition_seen
            monitor = self._query_monitor_state()
            transition_seen = transition_seen or bool(
                monitor["started"]
                or monitor["result_mutations"]
                or self._query_busy()
                or self._query_button_disabled()
                or self._slot_signature() != before
            )
            return transition_seen

        try:
            self.browser._click_element(query_button, "스케줄 조회")
        except StaleElementReferenceException:
            self.browser._click_element(self._find_query_button(), "스케줄 조회")

        self.browser._wait(
            30,
            query_transitioned,
            "일별 입고 조회 시작",
        )
        self.browser._wait(
            120,
            lambda: transition_seen
            and self._query_monitor_state()["active"] == 0
            and not self._query_busy()
            and not self._query_button_disabled(),
            "일별 입고 조회 완료",
        )
        failure = self._query_monitor_failure()
        if failure:
            raise DailyInboundError(f"일별 입고 조회 요청이 실패했습니다.\n{failure}")
        self.browser._wait(
            30,
            lambda: self._fresh_query_result_observed(before),
            "새 일별 입고 조회 결과 반영",
        )
        self.browser._wait(
            60,
            lambda: bool(
                self.browser._visible_elements(
                    By.CSS_SELECTOR,
                    "div.booking-slot-list mat-grid-list, div.booking-slot-list",
                )
            )
            or self._has_no_result_message(),
            "일별 입고 카드 목록",
        )

        previous = None
        stable_count = 0
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            self.browser._check_cancelled()
            signature = self._slot_signature()
            if signature == previous:
                stable_count += 1
            else:
                previous = signature
                stable_count = 1
            if stable_count >= 3:
                self._disarm_query_monitor()
                return
            self.browser.stop_event.wait(0.4)
        self._disarm_query_monitor()
        raise DailyInboundError(
            "일별 입고 카드 목록이 계속 변경되어 안정된 조회 결과를 확인하지 못했습니다."
        )

    def _matching_slots(self, dispatch_number: str) -> list[WebElement]:
        for _attempt in range(3):
            matches: list[WebElement] = []
            had_stale = False
            try:
                cards = self.browser._driver.find_elements(By.CSS_SELECTOR, "div.booking-slot")
            except WebDriverException:
                cards = []
            for card in cards:
                try:
                    labels = card.find_elements(By.CSS_SELECTOR, "b")
                    if (
                        labels
                        and normalize_milkrun_card_number(labels[0].text)
                        == dispatch_number
                    ):
                        matches.append(card)
                except (StaleElementReferenceException, WebDriverException):
                    had_stale = True
                    continue
            if matches or not had_stale:
                return matches
            self.browser.stop_event.wait(0.2)
        return []

    def _slot_signature(self) -> tuple[str, ...]:
        try:
            return tuple(
                normalize_milkrun_card_number(element.text)
                for element in self.browser._driver.find_elements(
                    By.CSS_SELECTOR,
                    "div.booking-slot b",
                )
                if normalize_milkrun_card_number(element.text)
            )
        except (StaleElementReferenceException, WebDriverException):
            return ()

    def _query_busy(self) -> bool:
        selectors = (
            "mat-progress-spinner",
            "mat-spinner",
            ".mat-mdc-progress-spinner",
            "[aria-busy='true']",
            ".loading:not([hidden])",
        )
        return any(self.browser._visible_elements(By.CSS_SELECTOR, selector) for selector in selectors)

    def _find_query_button(self) -> WebElement:
        selectors = (
            "//button[not(@disabled) and "
            "(normalize-space(.)='스케줄 조회' or .//*[normalize-space()='스케줄 조회'])]",
            "//button[not(@disabled) and "
            "(normalize-space(.)='조회' or .//*[normalize-space()='조회'])]",
        )
        last_error: TimeoutException | None = None
        for index, selector in enumerate(selectors):
            try:
                return self.browser._first_visible(
                    By.XPATH,
                    selector,
                    timeout=30 if index == 0 else 10,
                )
            except TimeoutException as exc:
                last_error = exc
        raise last_error or TimeoutException("스케줄 조회 버튼을 찾지 못했습니다.")

    def _query_button_disabled(self) -> bool:
        selectors = (
            "//button[normalize-space(.)='스케줄 조회' or .//*[normalize-space()='스케줄 조회']]",
            "//button[normalize-space(.)='조회' or .//*[normalize-space()='조회']]",
        )
        try:
            for selector in selectors:
                for button in self.browser._driver.find_elements(By.XPATH, selector):
                    if not button.is_displayed():
                        continue
                    return bool(
                        button.get_attribute("disabled") is not None
                        or button.get_attribute("aria-disabled") == "true"
                    )
        except (StaleElementReferenceException, WebDriverException):
            return False
        return False

    def _install_query_monitor(self) -> None:
        script = r"""
            window.__unhelperQueryState = {
                armed: false,
                started: 0,
                active: 0,
                completed: 0,
                successes: 0,
                failures: 0,
                querySuccesses: 0,
                queryFailures: 0,
                resultMutations: 0,
                lastQueryFinishedAt: 0,
                lastFailure: ''
            };
            if (window.__unhelperQueryHooksInstalled) return;
            window.__unhelperQueryHooksInstalled = true;

            const isQueryRequest = url =>
                /(?:inbound|schedule|booking|slot|milkrun|appointment|reservation)/i.test(url || '');
            const beginRequest = url => {
                const state = window.__unhelperQueryState;
                if (!state || !state.armed) return null;
                const record = {url: String(url || ''), query: isQueryRequest(String(url || ''))};
                state.started += 1;
                state.active += 1;
                return {state, record};
            };
            const finishRequest = (context, ok, status, error) => {
                if (!context) return;
                const {state, record} = context;
                state.active = Math.max(0, state.active - 1);
                state.completed += 1;
                if (ok) state.successes += 1;
                else {
                    state.failures += 1;
                    state.lastFailure = `${record.url || '(URL 없음)'} · ${status || error || 'network error'}`;
                }
                if (record.query) {
                    if (ok) state.querySuccesses += 1;
                    else state.queryFailures += 1;
                    state.lastQueryFinishedAt = performance.now();
                }
            };

            const relatedToResult = node => {
                const element = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
                if (!element) return false;
                return Boolean(
                    element.matches?.('div.booking-slot-list, div.booking-slot, mat-grid-list') ||
                    element.closest?.('div.booking-slot-list') ||
                    element.querySelector?.('div.booking-slot-list, div.booking-slot, mat-grid-list')
                );
            };
            new MutationObserver(records => {
                const state = window.__unhelperQueryState;
                if (!state || !state.armed) return;
                if (records.some(record =>
                    relatedToResult(record.target) ||
                    Array.from(record.addedNodes).some(relatedToResult) ||
                    Array.from(record.removedNodes).some(relatedToResult)
                )) state.resultMutations += 1;
            }).observe(document.body, {
                subtree: true,
                childList: true,
                characterData: true,
                attributes: true
            });

            const originalFetch = window.fetch;
            if (typeof originalFetch === 'function') {
                window.fetch = function(...args) {
                    const input = args[0];
                    const url = typeof input === 'string' ? input : (input && input.url) || '';
                    const context = beginRequest(url);
                    let result;
                    try {
                        result = originalFetch.apply(this, args);
                    } catch (error) {
                        finishRequest(context, false, 0, String(error));
                        throw error;
                    }
                    return Promise.resolve(result).then(response => {
                        const status = Number(response.status || 0);
                        finishRequest(context, status >= 200 && status < 400, status, '');
                        return response;
                    }, error => {
                        finishRequest(context, false, 0, String(error));
                        throw error;
                    });
                };
            }

            const originalSend = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.send = function(...args) {
                const context = beginRequest(this.__unhelperUrl || '');
                this.addEventListener('loadend', () => {
                    const status = Number(this.status || 0);
                    if (context) context.record.url = this.responseURL || context.record.url;
                    finishRequest(context, status >= 200 && status < 400, status, '');
                }, {once: true});
                try {
                    return originalSend.apply(this, args);
                } catch (error) {
                    finishRequest(context, false, 0, String(error));
                    throw error;
                }
            };
            const originalOpen = XMLHttpRequest.prototype.open;
            XMLHttpRequest.prototype.open = function(method, url, ...rest) {
                this.__unhelperUrl = String(url || '');
                return originalOpen.call(this, method, url, ...rest);
            };
        """
        try:
            self.browser._driver.execute_script(script)
        except WebDriverException:
            # The visible busy/button/result transitions remain valid fallbacks.
            pass

    def _arm_query_monitor(self) -> None:
        try:
            self.browser._driver.execute_script(
                "if (window.__unhelperQueryState) { "
                "window.__unhelperQueryState.armed = true; "
                "window.__unhelperQueryState.clickAt = performance.now(); }"
            )
        except WebDriverException:
            pass

    def _disarm_query_monitor(self) -> None:
        try:
            self.browser._driver.execute_script(
                "if (window.__unhelperQueryState) window.__unhelperQueryState.armed = false;"
            )
        except WebDriverException:
            pass

    def _query_monitor_state(self) -> dict[str, int | str]:
        try:
            state = self.browser._driver.execute_script(
                "const state = window.__unhelperQueryState || {}; "
                "return {...state, now: performance.now()};"
            ) or {}
            finished_at = float(state.get("lastQueryFinishedAt", 0) or 0)
            now = float(state.get("now", 0) or 0)
            return {
                "started": int(state.get("started", 0) or 0),
                "active": int(state.get("active", 0) or 0),
                "successes": int(state.get("successes", 0) or 0),
                "failures": int(state.get("failures", 0) or 0),
                "query_successes": int(state.get("querySuccesses", 0) or 0),
                "query_failures": int(state.get("queryFailures", 0) or 0),
                "result_mutations": int(state.get("resultMutations", 0) or 0),
                "query_settled_ms": int(max(0, now - finished_at)) if finished_at else 0,
                "last_failure": str(state.get("lastFailure", "") or ""),
            }
        except (TypeError, ValueError, WebDriverException):
            return {
                "started": 0,
                "active": 0,
                "successes": 0,
                "failures": 0,
                "query_successes": 0,
                "query_failures": 0,
                "result_mutations": 0,
                "query_settled_ms": 0,
                "last_failure": "",
            }

    def _query_monitor_failure(self) -> str:
        state = self._query_monitor_state()
        if int(state["query_failures"]) > 0:
            return str(state["last_failure"]) or "조회 관련 요청이 HTTP 오류로 종료되었습니다."
        return ""

    def _fresh_query_result_observed(self, before: tuple[str, ...]) -> bool:
        state = self._query_monitor_state()
        if int(state["result_mutations"]) > 0:
            return True
        if self._slot_signature() != before or self._has_no_result_message():
            return True
        return bool(
            int(state["query_successes"]) > 0
            and int(state["query_settled_ms"]) >= 1000
        )

    def _has_no_result_message(self) -> bool:
        try:
            text = self.browser._normalize(self.browser._driver.find_element(By.TAG_NAME, "body").text)
        except WebDriverException:
            return False
        return any(marker in text for marker in ("조회 결과가 없습니다", "데이터가 없습니다", "No Data"))

    def _open_detail_and_read(
        self,
        card: WebElement,
        dispatch_number: str,
    ) -> tuple[MilkrunProductRow, ...]:
        original_handle = self.browser._driver.current_window_handle
        handles_before = set(self.browser._driver.window_handles)
        created_handles: set[str] = set()
        try:
            self.browser._click_element(card, f"배차번호 {dispatch_number} 카드")
            link = self.browser._first_visible(
                By.XPATH,
                "//a[@target='_blank' and contains(@href,'/app/inbound-booking/milkrun/detail')]"
                "[.//em[normalize-space()='선택한 Slot의 예약정보를 조회합니다.']]",
                timeout=20,
            )
            href = link.get_attribute("href") or ""
            self.browser._click_element(link, "선택한 Slot의 예약정보 조회")
            try:
                created_handles = set(
                    self.browser._wait(
                        12,
                        lambda: set(self.browser._driver.window_handles) - handles_before,
                        "예약 상세 새 창",
                    )
                )
            except TimeoutException:
                created_handles = set(self.browser._driver.window_handles) - handles_before
                if not created_handles and href:
                    self.browser._driver.execute_script("window.open(arguments[0], '_blank');", href)
                    created_handles = set(
                        self.browser._wait(
                            12,
                            lambda: set(self.browser._driver.window_handles) - handles_before,
                            "예약 상세 새 창",
                        )
                    )
            if not created_handles:
                raise DailyInboundError("예약 상세 페이지가 새 창으로 열리지 않았습니다.")

            detail_handle = next(iter(created_handles))
            self.browser._driver.switch_to.window(detail_handle)
            self.browser._wait_document_ready(60)
            self.browser._wait(
                60,
                lambda: self.DETAIL_HREF_FRAGMENT in self._safe_current_url(),
                "밀크런 예약 상세 화면",
            )
            logical_rows = self._wait_for_detail_rows()
            products = parse_detail_table_cells(
                logical_rows,
                dispatch_number=dispatch_number,
            )
            if not products:
                raise DailyInboundError("예약 상세 표의 SKU 열을 읽지 못했습니다. 사이트 표 구조를 확인해 주세요.")
            return products
        except Exception as exc:
            if not self.browser.stop_event.is_set() and self.evidence_dir is not None:
                failure_url = self._safe_current_url()
                self.browser.save_failure_snapshot(self.evidence_dir, exc)
                raise DailyInboundError(
                    f"{exc}\n실패 상세 주소: {failure_url}",
                    evidence_captured=True,
                    failure_url=failure_url,
                ) from exc
            raise
        finally:
            pending_exception = sys.exc_info()[0] is not None
            try:
                open_handles = set(self.browser._driver.window_handles)
                for handle in created_handles & open_handles:
                    try:
                        self.browser._driver.switch_to.window(handle)
                        self.browser._driver.close()
                    except (NoSuchWindowException, WebDriverException):
                        pass
                if original_handle in self.browser._driver.window_handles:
                    self.browser._driver.switch_to.window(original_handle)
                    try:
                        self._dismiss_detail_sheet()
                    except DailyInboundError as cleanup_error:
                        if not pending_exception:
                            raise
                        self.log(f"예약 상세 선택창 정리 경고: {cleanup_error}")
            except (NoSuchWindowException, WebDriverException):
                pass

    def _wait_for_detail_rows(self) -> tuple[tuple[str, ...], ...]:
        previous: tuple[tuple[str, ...], ...] | None = None
        stable_count = 0
        deadline = time.monotonic() + 90
        while time.monotonic() < deadline:
            self.browser._check_cancelled()
            if not self._query_busy():
                rows = self._detail_logical_rows()
                if rows:
                    if rows == previous:
                        stable_count += 1
                    else:
                        previous = rows
                        stable_count = 1
                    if stable_count >= 3:
                        return rows
                elif self._has_no_result_message():
                    raise DailyInboundError("예약 상세 페이지에 상품 데이터가 없습니다.")
            self.browser.stop_event.wait(0.4)
        raise TimeoutException("예약 상세 상품 표를 90초 안에 읽지 못했습니다.")

    def _detail_logical_rows(self) -> tuple[tuple[str, ...], ...]:
        script = r"""
            const normalize = value => (value || '').replace(/\s+/g, ' ').trim();
            const expand = body => {
                const carries = [];
                const result = [];
                for (const tr of Array.from(body.querySelectorAll(':scope > tr'))) {
                    const logical = [];
                    for (let col = 0; col < carries.length; col += 1) {
                        const carry = carries[col];
                        if (!carry || carry.remaining <= 0) continue;
                        logical[col] = carry.text;
                        carry.remaining -= 1;
                        if (carry.remaining <= 0) carries[col] = null;
                    }
                    let col = 0;
                    for (const cell of Array.from(tr.children).filter(node => node.tagName === 'TD')) {
                        while (logical[col] !== undefined) col += 1;
                        const text = normalize(cell.innerText);
                        const rowSpan = Math.max(1, Number.parseInt(cell.getAttribute('rowspan') || '1', 10));
                        const colSpan = Math.max(1, Number.parseInt(cell.getAttribute('colspan') || '1', 10));
                        for (let offset = 0; offset < colSpan; offset += 1) {
                            const value = offset === 0 ? text : '';
                            logical[col + offset] = value;
                            if (rowSpan > 1) {
                                carries[col + offset] = {text: value, remaining: rowSpan - 1};
                            }
                        }
                        col += colSpan;
                    }
                    result.push(logical.map(value => value === undefined ? '' : value));
                }
                return result;
            };
            const candidates = Array.from(document.querySelectorAll('table tbody'))
                .map(expand)
                .filter(rows => rows.some(row => row.length >= 8 && row[6] && row[7]));
            candidates.sort((left, right) => right.length - left.length);
            return candidates.length ? candidates[0] : [];
        """
        try:
            raw_rows = self.browser._driver.execute_script(script) or []
        except WebDriverException:
            return ()
        return tuple(tuple(str(value or "") for value in row) for row in raw_rows)

    def _dismiss_detail_sheet(self) -> None:
        try:
            ActionChains(self.browser._driver).send_keys(Keys.ESCAPE).perform()
        except WebDriverException:
            pass
        try:
            self.browser._wait(
                5,
                self._detail_sheet_closed,
                "예약 상세 선택창 닫기",
            )
        except TimeoutException:
            backdrops = self.browser._visible_elements(By.CSS_SELECTOR, ".cdk-overlay-backdrop")
            if backdrops:
                try:
                    self.browser._click_element(backdrops[-1], "예약 상세 선택창 배경")
                except RuntimeError:
                    pass
            try:
                self.browser._wait(
                    5,
                    self._detail_sheet_closed,
                    "예약 상세 선택창 닫힘",
                )
            except TimeoutException as exc:
                raise DailyInboundError("예약 상세 선택창을 닫지 못했습니다.") from exc

    def _detail_sheet_closed(self) -> bool:
        detail_links = self.browser._visible_elements(
            By.XPATH,
            "//a[contains(@href,'/app/inbound-booking/milkrun/detail')]"
            "[.//em[normalize-space()='선택한 Slot의 예약정보를 조회합니다.']]",
        )
        backdrops = self.browser._visible_elements(By.CSS_SELECTOR, ".cdk-overlay-backdrop")
        return not detail_links and not backdrops

    def _safe_current_url(self) -> str:
        try:
            return self.browser._driver.current_url
        except WebDriverException:
            return "(주소 확인 불가)"
