"""Lectura de las páginas reales del SII (capturadas el 22-10-2026) y reglas de tasas."""
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path
from unittest import mock

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout

from rpa_andes import sii
from rpa_andes.sii import (ExtractionError, RateUnavailable, StructureError, parse_rates, to_decimal,
                           uf_for, usd_for)

FIXTURES = Path(__file__).parent / "fixtures"
UF_HTML = (FIXTURES / "uf2026.html").read_bytes()
USD_HTML = (FIXTURES / "dolar2026.html").read_bytes()


class ParseTests(unittest.TestCase):
    def test_uf_annual_table(self):
        values = parse_rates(UF_HTML, "UF", 2026)
        self.assertEqual(len(values), 313)
        self.assertEqual(values[date(2026, 1, 1)], "39.731,79")
        self.assertEqual(max(values), date(2026, 11, 9))

    def test_usd_annual_table_ignores_average_row(self):
        values = parse_rates(USD_HTML, "USD", 2026)
        self.assertEqual(len(values), 189)
        self.assertEqual(to_decimal(values[date(2026, 10, 1)]), Decimal("972.60"))
        self.assertNotIn(date(2026, 9, 5), values)  # sábado

    def test_wrong_page_or_year_fails(self):
        with self.assertRaisesRegex(StructureError, "no corresponde a USD"):
            parse_rates(UF_HTML, "USD", 2026)
        with self.assertRaisesRegex(StructureError, "año 2025"):
            parse_rates(UF_HTML, "UF", 2025)

    def test_changed_header_fails(self):
        html = USD_HTML.replace(b"<th>Oct</th>", b"<th>Octubre</th>")
        with self.assertRaisesRegex(StructureError, "cabecera"):
            parse_rates(html, "USD", 2026)

    def test_changed_number_format_fails(self):
        html = USD_HTML.replace(b">972,60<", b">972.60<", 1)
        with self.assertRaisesRegex(StructureError, "Formato de valor"):
            parse_rates(html, "USD", 2026)

    def test_missing_table_fails(self):
        html = USD_HTML.replace(b'id="table_export"', b"").replace(b"<th>D&iacute;a</th>", b"<th>X</th>")
        with self.assertRaisesRegex(StructureError, "encontraron 0"):
            parse_rates(html, "USD", 2026)

    def test_missing_day_row_fails(self):
        html = USD_HTML.replace(b"<th id='f15' style='text-align:center;'>15</th>",
                                b"<th id='f15' style='text-align:center;'>16</th>")
        with self.assertRaisesRegex(StructureError, "Fila inesperada"):
            parse_rates(html, "USD", 2026)


class RateRuleTests(unittest.TestCase):
    usd = {d: to_decimal(v) for d, v in parse_rates(USD_HTML, "USD", 2026).items()}
    uf = {d: to_decimal(v) for d, v in parse_rates(UF_HTML, "UF", 2026).items()}

    def test_published_day_uses_its_own_value(self):
        self.assertEqual(usd_for(date(2026, 9, 1), self.usd, 7), (date(2026, 9, 1), Decimal("933.40"), 0))

    def test_saturday_uses_friday(self):
        self.assertEqual(usd_for(date(2026, 9, 5), self.usd, 7)[0], date(2026, 9, 4))

    def test_fiestas_patrias_use_last_published_day(self):
        used, _, age = usd_for(date(2026, 9, 19), self.usd, 7)
        self.assertEqual((used, age), (date(2026, 9, 17), 2))

    def test_business_day_after_last_publication_is_not_guessed(self):
        # El SII publicó hasta el jueves 01-10; el viernes 02-10 aún no aparece.
        for day in (date(2026, 10, 2), date(2026, 10, 3)):
            with self.assertRaisesRegex(RateUnavailable, "aún no publica el dólar del 02-10-2026"):
                usd_for(day, self.usd, 7)

    def test_weekend_after_last_publication_is_carried(self):
        series = {d: v for d, v in self.usd.items() if d <= date(2026, 9, 4)}  # termina un viernes
        used, _, age = usd_for(date(2026, 9, 6), series, 7)
        self.assertEqual((used, age), (date(2026, 9, 4), 2))

    def test_carry_never_exceeds_the_limit(self):
        series = {date(2026, 9, 1): Decimal("900"), date(2026, 9, 30): Decimal("950")}
        with self.assertRaisesRegex(RateUnavailable, "máximo 7"):
            usd_for(date(2026, 9, 10), series, 7)

    def test_uf_requires_exact_date(self):
        self.assertEqual(uf_for(date(2026, 9, 19), self.uf)[1], Decimal("40967.24"))
        with self.assertRaises(RateUnavailable):
            uf_for(date(2026, 12, 1), self.uf)


class FakeResponse:
    def __init__(self, status, url="https://www.sii.cl/x.htm"):
        self.status, self.url = status, url

    def body(self):
        return b"<html></html>"


class FetchTests(unittest.TestCase):
    """Reintentos del navegador, con Playwright simulado."""
    browser = {"canal": "", "headless": True, "timeout_ms": 1000, "intentos": 3, "espera_inicial_s": 0}

    def fetch(self, *outcomes):
        page = mock.Mock()
        page.goto.side_effect = list(outcomes)
        playwright = mock.MagicMock()
        playwright.__enter__.return_value.chromium.launch.return_value.new_page.return_value = page
        with mock.patch.object(sii, "sync_playwright", return_value=playwright):
            try:
                return sii.fetch_page("https://www.sii.cl/x.htm", self.browser, mock.Mock()), page.goto.call_count
            except ExtractionError as exc:
                return exc, page.goto.call_count

    def test_server_error_then_success(self):
        result, calls = self.fetch(FakeResponse(503), FakeResponse(200))
        self.assertEqual((result["http"], calls), (200, 2))

    def test_network_and_proxy_errors_are_retried(self):
        result, calls = self.fetch(PlaywrightError("net::ERR_TUNNEL_CONNECTION_FAILED"),
                                   PlaywrightError("net::ERR_CONNECTION_RESET"), FakeResponse(200))
        self.assertEqual((result["http"], calls), (200, 3))

    def test_gives_up_after_configured_attempts(self):
        result, calls = self.fetch(*[PlaywrightTimeout("timeout")] * 3)
        self.assertIsInstance(result, ExtractionError)
        self.assertEqual(calls, 3)

    def test_client_error_and_certificate_error_are_not_retried(self):
        for outcome in (FakeResponse(404), PlaywrightError("net::ERR_CERT_AUTHORITY_INVALID")):
            result, calls = self.fetch(outcome, FakeResponse(200))
            self.assertIsInstance(result, ExtractionError)
            self.assertEqual(calls, 1)

    def test_redirect_outside_sii_is_rejected(self):
        result, _ = self.fetch(FakeResponse(200, url="https://otro-sitio.cl/"))
        self.assertIsInstance(result, ExtractionError)


if __name__ == "__main__":
    unittest.main()
