from __future__ import annotations

import threading
import unittest
from decimal import Decimal
from pathlib import Path
from typing import Callable
from unittest import mock

from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By

from Modules.WMS.ProductWeightCrawler import (
    ProductWeightCrawler,
    WMSWeightError,
    _DetailContext,
    _SearchResult,
)


class _Element:
    def __init__(
        self,
        text: str = "",
        *,
        value: str = "",
        href: str = "",
        target: str = "",
        links: list[object] | None = None,
        columns: list[object] | None = None,
        stale: bool = False,
        on_click: Callable[[], None] | None = None,
    ) -> None:
        self.text = text
        self.value = value
        self.href = href
        self.target = target
        self.links = links or []
        self.columns = columns or []
        self.stale = stale
        self.on_click = on_click

    def click(self) -> None:
        if self.stale:
            raise StaleElementReferenceException("rerendered")
        if self.on_click is not None:
            self.on_click()

    def find_elements(self, by, selector):
        if self.stale:
            raise StaleElementReferenceException("rerendered")
        if (by, selector) == (By.TAG_NAME, "td"):
            return list(self.columns)
        if (by, selector) == (By.TAG_NAME, "a"):
            return list(self.links)
        return []

    def get_attribute(self, name: str):
        if self.stale:
            raise StaleElementReferenceException("rerendered")
        if name == "value":
            return self.value
        if name == "href":
            return self.href
        if name == "target":
            return self.target
        return ""

    def is_enabled(self) -> bool:
        if self.stale:
            raise StaleElementReferenceException("rerendered")
        return True

    def is_displayed(self) -> bool:
        return True


def _result_row(
    sku: str,
    product_name: str,
    href: str = "/sku/detail",
    *,
    target: str = "",
    column_count: int = 11,
) -> _Element:
    link = _Element(product_name, href=href, target=target)
    columns = [_Element() for _ in range(column_count)]
    columns[0] = _Element(sku)
    columns[10] = _Element(product_name, links=[link])
    return _Element(columns=columns)


class _SwitchTo:
    def __init__(self, driver: "_TabDriver") -> None:
        self.driver = driver

    def window(self, handle: str) -> None:
        self.driver.current_window_handle = handle


class _TabDriver:
    def __init__(
        self,
        handles: list[str],
        current: str,
        current_url: str = "https://wms.coupang.com/sku/list",
    ) -> None:
        self.window_handles = list(handles)
        self.current_window_handle = current
        self.current_url = current_url
        self.switch_to = _SwitchTo(self)
        self.closed: list[str] = []
        self.back_calls = 0

    def close(self) -> None:
        self.closed.append(self.current_window_handle)
        self.window_handles.remove(self.current_window_handle)

    def back(self) -> None:
        self.back_calls += 1


class ProductWeightCrawlerTests(unittest.TestCase):
    def make_crawler(self) -> ProductWeightCrawler:
        return ProductWeightCrawler(Path("chromedriver.exe"), log=lambda _message: None)

    def test_sku_input_targets_product_management_external_id_field(self) -> None:
        self.assertEqual(
            ProductWeightCrawler.SKU_INPUT_LOCATOR,
            (By.CSS_SELECTOR, "input.form-control.input-external-id"),
        )

    def test_exact_sku_row_only_and_product_name_preserves_slash(self) -> None:
        wrong = _result_row("999", "다른 상품")
        exact = _result_row("123", "상품 A /\n상품 B")

        self.assertIsNone(ProductWeightCrawler._parse_result_row(wrong, "123"))
        parsed = ProductWeightCrawler._parse_result_row(exact, "123")

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.sku_id, "123")
        self.assertEqual(parsed.product_name, "상품 A / 상품 B")
        self.assertIn("/", parsed.product_name)
        self.assertNotIn("\n", parsed.product_name)

    def test_actual_wms_row_uses_eleventh_column_link_without_relying_on_target(self) -> None:
        row = _result_row(
            "163108821",
            "Box*크리넥스 안심 다용도 타월 100매X10개 , 1박스",
            "/sku/163108821",
            target="*blank",
            column_count=16,
        )

        parsed = ProductWeightCrawler._parse_result_row(row, "163108821")

        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.sku_id, "163108821")
        self.assertEqual(parsed.product_name, "Box*크리넥스 안심 다용도 타월 100매X10개 , 1박스")
        self.assertEqual(parsed.href, "/sku/163108821")
        self.assertEqual(parsed.link.get_attribute("target"), "*blank")

    def test_exact_sku_must_come_from_first_column_even_when_link_points_to_it(self) -> None:
        row = _result_row(
            "999999999",
            "다른 SKU 행",
            "/sku/163108821",
            target="_blank",
            column_count=16,
        )

        self.assertIsNone(ProductWeightCrawler._parse_result_row(row, "163108821"))

    def test_sku_normalization_removes_excel_decimal_but_keeps_leading_zero(self) -> None:
        self.assertEqual(ProductWeightCrawler.normalize_sku("1,234.0"), "1234")
        self.assertEqual(ProductWeightCrawler.normalize_sku("00123"), "00123")
        self.assertEqual(ProductWeightCrawler.normalize_sku(" 12 34 "), "1234")

    def test_empty_weight_is_an_error(self) -> None:
        crawler = self.make_crawler()
        with mock.patch.object(crawler, "_first_present", return_value=_Element(value="")):
            with self.assertRaisesRegex(WMSWeightError, "비어"):
                crawler._extract_weight_grams()

    def test_zero_and_negative_weight_are_errors(self) -> None:
        crawler = self.make_crawler()
        for raw_value in ("0", "-1"):
            with self.subTest(raw_value=raw_value):
                with mock.patch.object(
                    crawler,
                    "_first_present",
                    return_value=_Element(value=raw_value),
                ):
                    with self.assertRaisesRegex(WMSWeightError, "0 이하"):
                        crawler._extract_weight_grams()

    def test_positive_weight_is_decimal_grams(self) -> None:
        crawler = self.make_crawler()
        with mock.patch.object(crawler, "_first_present", return_value=_Element(value="1,250.5")):
            self.assertEqual(crawler._extract_weight_grams(), Decimal("1250.5"))

    def test_actual_hidden_weight_value_is_read_as_grams(self) -> None:
        crawler = self.make_crawler()
        with mock.patch.object(crawler, "_first_present", return_value=_Element(value="1450")) as find:
            self.assertEqual(crawler._extract_weight_grams(), Decimal("1450"))

        find.assert_called_once_with(
            By.CSS_SELECTOR,
            "input.hidden-weight",
            timeout=crawler.timeout,
        )

    def test_open_product_detail_detects_new_tab_without_reading_target_attribute(self) -> None:
        crawler = self.make_crawler()
        driver = _TabDriver(["main"], "main")
        crawler.driver = driver

        def open_detail_tab() -> None:
            driver.window_handles.append("detail")

        link = _Element(
            "상품명",
            href="/sku/163108821",
            target="*blank",
            on_click=open_detail_tab,
        )

        with mock.patch.object(crawler, "_wait_document_ready"):
            context = crawler._open_product_detail(link)

        self.assertTrue(context.opened_new_tab)
        self.assertEqual(context.original_handle, "main")
        self.assertEqual(context.detail_handle, "detail")
        self.assertEqual(driver.current_window_handle, "detail")

    def test_open_product_detail_detects_same_tab_navigation(self) -> None:
        crawler = self.make_crawler()
        driver = _TabDriver(["main"], "main")
        crawler.driver = driver
        link = _Element(
            "상품명",
            href="/sku/163108821",
            on_click=lambda: setattr(
                driver,
                "current_url",
                "https://wms.coupang.com/sku/163108821",
            ),
        )

        with mock.patch.object(crawler, "_wait_document_ready"):
            context = crawler._open_product_detail(link)

        self.assertFalse(context.opened_new_tab)
        self.assertEqual(context.original_handle, "main")
        self.assertEqual(context.detail_handle, "main")
        self.assertEqual(driver.current_window_handle, "main")

    def test_new_detail_tab_is_closed_and_original_tab_is_restored(self) -> None:
        crawler = self.make_crawler()
        driver = _TabDriver(["main", "detail"], "detail")
        crawler.driver = driver
        context = _DetailContext("main", "detail", True, "https://wms/product")

        crawler._restore_after_detail(context)

        self.assertEqual(driver.closed, ["detail"])
        self.assertEqual(driver.window_handles, ["main"])
        self.assertEqual(driver.current_window_handle, "main")
        self.assertEqual(driver.back_calls, 0)

    def test_same_detail_tab_uses_back_without_closing(self) -> None:
        crawler = self.make_crawler()
        driver = _TabDriver(["main"], "main")
        crawler.driver = driver
        context = _DetailContext("main", "main", False, "https://wms/product")

        with mock.patch.object(crawler, "_wait_document_ready"):
            crawler._restore_after_detail(context)

        self.assertEqual(driver.closed, [])
        self.assertEqual(driver.back_calls, 1)
        self.assertEqual(driver.current_window_handle, "main")

    def test_fresh_result_must_transition_and_stabilize(self) -> None:
        crawler = self.make_crawler()
        crawler.timeout = 0.05
        crawler.SKU_SEARCH_TIMEOUT_SECONDS = 0.05
        crawler.POLL_INTERVAL_SECONDS = 0.001
        link = _Element(href="/sku/detail/123")
        candidate = _SearchResult("123", "상품 / 이름", link, link.href)

        with (
            mock.patch.object(crawler, "_ensure_browser_open"),
            mock.patch.object(crawler, "_table_signature", return_value=("new",)),
            mock.patch.object(crawler, "_is_stale", return_value=False),
            mock.patch.object(crawler, "_result_mutation_count", return_value=0),
            mock.patch.object(crawler, "_find_exact_result", return_value=candidate),
            mock.patch.object(crawler, "_has_no_result_message", return_value=False),
        ):
            result = crawler._wait_for_fresh_result("123", ("old",), None)

        self.assertIs(result, candidate)

    def test_old_exact_row_is_not_accepted_without_fresh_transition(self) -> None:
        crawler = self.make_crawler()
        crawler.timeout = 0.01
        crawler.SKU_SEARCH_TIMEOUT_SECONDS = 0.01
        crawler.POLL_INTERVAL_SECONDS = 0.001

        with (
            mock.patch.object(crawler, "_ensure_browser_open"),
            mock.patch.object(crawler, "_table_signature", return_value=("old",)),
            mock.patch.object(crawler, "_is_stale", return_value=False),
            mock.patch.object(crawler, "_result_mutation_count", return_value=0),
            mock.patch.object(crawler, "_find_exact_result") as find_exact,
        ):
            with self.assertRaisesRegex(WMSWeightError, "새 검색 결과"):
                crawler._wait_for_fresh_result("123", ("old",), None)

        find_exact.assert_not_called()

    def test_cancel_interrupts_waits(self) -> None:
        stop_event = threading.Event()
        stop_event.set()
        crawler = ProductWeightCrawler(
            Path("chromedriver.exe"),
            stop_event=stop_event,
            log=lambda _message: None,
        )

        with self.assertRaisesRegex(WMSWeightError, "중지"):
            crawler._check_cancelled()

    def test_html_redaction_masks_login_fields_and_known_credentials(self) -> None:
        crawler = ProductWeightCrawler(
            Path("chromedriver.exe"),
            user_id="worker123",
            user_password="plain-secret",
            log=lambda _message: None,
        )
        html = (
            '<input id="username" value="worker123">'
            '<input type="password" value="plain-secret">'
            '<input id=password value=unquoted-secret>'
            '<script>const state={"token":"abc123"};</script>'
            'Authorization: Bearer bearer-secret'
        )

        redacted = crawler.redact_sensitive_html(html)

        self.assertNotIn("worker123", redacted)
        self.assertNotIn("plain-secret", redacted)
        self.assertNotIn("unquoted-secret", redacted)
        self.assertNotIn("abc123", redacted)
        self.assertNotIn("bearer-secret", redacted)
        self.assertGreaterEqual(redacted.count("***"), 5)


if __name__ == "__main__":
    unittest.main()
