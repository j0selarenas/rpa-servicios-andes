"""Extracción de UF y dólar observado desde el sitio del SII.

El SII no tiene API. Playwright abre la página anual con un navegador real y se
guarda la respuesta HTML original en landing. Después se lee solo la tabla
anual (Día x Ene..Dic); las tablas mensuales repiten los mismos valores.

Cualquier diferencia con la estructura esperada detiene el proceso: es
preferible no cargar nada a cargar un valor mal leído.
"""
from __future__ import annotations

import re
import time
import unicodedata
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal

from bs4 import BeautifulSoup
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeout
from playwright.sync_api import sync_playwright

MONTHS = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")
VALUE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")


class ExtractionError(RuntimeError):
    """Error de sistema: el sitio no respondió o respondió con error."""


class StructureError(RuntimeError):
    """Error de sistema: la página cambió y no se puede leer con seguridad."""


class RateUnavailable(Exception):
    """Excepción de negocio: no hay un valor válido para la fecha de la solicitud."""


def fetch_page(url: str, browser: dict, log) -> dict:
    """Abre la URL en el navegador y devuelve el HTML original de la respuesta.

    Se reintenta ante timeout, errores de red (net::ERR_*) y HTTP 408/429/5xx.
    No se reintenta un error de certificado ni un HTTP 4xx: repetir no lo arregla.
    """
    attempts = browser["intentos"]
    options = {"headless": browser["headless"]}
    if browser["canal"]:
        options["channel"] = browser["canal"]

    with sync_playwright() as pw:
        try:
            instance = pw.chromium.launch(**options)
        except PlaywrightError as exc:
            raise ExtractionError(f"No se pudo abrir el navegador ({browser['canal'] or 'chromium'}): {exc}") from exc
        try:
            page = instance.new_page()
            for attempt in range(1, attempts + 1):
                try:
                    response = page.goto(url, wait_until="domcontentloaded", timeout=browser["timeout_ms"])
                    status = response.status if response else None
                    if status == 200:
                        if not response.url.startswith("https://www.sii.cl/"):
                            raise ExtractionError(f"El SII redirigió fuera del sitio: {response.url}")
                        return {
                            "html": response.body(),
                            "url": response.url,
                            "http": status,
                            "capturado_en": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                        }
                    if status is not None and status < 500 and status not in (408, 429):
                        raise ExtractionError(f"El SII respondió HTTP {status} para {url}")
                    reason = f"HTTP {status}"
                except PlaywrightTimeout:
                    reason = f"sin respuesta en {browser['timeout_ms'] // 1000} s"
                except PlaywrightError as exc:
                    message = str(exc).splitlines()[0]
                    if "net::ERR_" not in message or "ERR_CERT" in message:
                        raise ExtractionError(f"Error del navegador: {message}") from exc
                    reason = message

                if attempt == attempts:
                    raise ExtractionError(f"El SII no respondió después de {attempts} intentos ({reason})")
                wait = browser["espera_inicial_s"] * 2 ** (attempt - 1)
                log.warning("Intento %s/%s falló (%s). Reintento en %s s", attempt, attempts, reason, wait)
                time.sleep(wait)
        finally:
            instance.close()
    raise ExtractionError("No se obtuvo respuesta del SII")


def _clean(text: str) -> str:
    text = " ".join(text.replace("\xa0", " ").split()).lower()
    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def parse_rates(html: bytes, indicator: str, year: int) -> dict[date, str]:
    """Lee la tabla anual de UF o dólar y devuelve {fecha: texto de la celda}, p. ej. "39.731,79".

    Celdas vacías = día sin publicación (o que no existe, como el 30 de febrero).
    La fila "Promedio" del dólar no es un valor diario y se ignora.
    """
    soup = BeautifulSoup(html, "html.parser")
    # El selector de año lista todos los años; no sirve para saber cuál es esta página.
    for tag in soup(["script", "style", "select", "option"]):
        tag.decompose()
    headings = _clean(" ".join(h.get_text(" ") for h in soup.find_all(["h1", "h2", "h3", "h4"])))
    expected = "uf" if indicator == "UF" else "dolar observado"
    if not re.search(rf"\b{expected}\b", headings):
        raise StructureError(f"La página no corresponde a {indicator} (títulos: {headings[:80]!r})")
    if not re.search(rf"\bano {year}\b", _clean(soup.get_text(" "))):
        raise StructureError(f"La página de {indicator} no corresponde al año {year}")

    header = ("dia",) + MONTHS
    tables = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        first = tuple(_clean(c.get_text()) for c in rows[0].find_all(["th", "td"])) if rows else ()
        if first[:1] == ("dia",) and len(first) == 13:
            if first != header:
                raise StructureError(f"Cambió la cabecera de la tabla anual de {indicator}: {first}")
            tables.append(rows)
    if len(tables) != 1:
        raise StructureError(f"Se esperaba una tabla anual de {indicator} y se encontraron {len(tables)}")

    values: dict[date, str] = {}
    day_expected = 1
    for row in tables[0][1:]:
        cells = [c.get_text(" ", strip=True) for c in row.find_all(["th", "td"], recursive=False)]
        label = _clean(cells[0]) if cells else ""
        if label == "promedio" and indicator == "USD":
            continue
        if len(cells) != 13 or label != str(day_expected):
            raise StructureError(f"Fila inesperada en la tabla de {indicator}: {cells[:3]}")
        for month, raw in enumerate(cells[1:], start=1):
            raw = raw.replace("\xa0", "").strip()
            if not raw:
                continue
            if not VALUE.fullmatch(raw):
                raise StructureError(f"Formato de valor inesperado en {indicator} día {day_expected}/{month}: {raw!r}")
            try:
                day = date(year, month, day_expected)
            except ValueError:
                raise StructureError(f"Valor en una fecha que no existe: {day_expected}/{month}/{year}") from None
            values[day] = raw
        day_expected += 1

    if day_expected != 32:
        raise StructureError(f"La tabla de {indicator} no tiene los 31 días")
    if not values:
        raise StructureError(f"La tabla de {indicator} {year} no tiene valores")
    return values


def to_decimal(raw: str) -> Decimal:
    """'39.731,79' -> Decimal('39731.79'). El formato ya fue validado en parse_rates."""
    return Decimal(raw.replace(".", "").replace(",", "."))


def uf_for(day: date, series: dict[date, Decimal]) -> tuple[date, Decimal]:
    """La UF se publica para todos los días, incluso feriados: se exige la fecha exacta."""
    if day not in series:
        raise RateUnavailable(f"no hay UF publicada para el {day:%d-%m-%Y}")
    return day, series[day]


def usd_for(day: date, series: dict[date, Decimal], max_days: int) -> tuple[date, Decimal, int]:
    """Dólar observado vigente para una fecha. Devuelve (fecha usada, valor, días de arrastre).

    1. Si hay valor publicado ese día, se usa ese.
    2. Si no (fin de semana o feriado), se usa el último publicado antes, con un
       máximo de `max_days` días corridos.
    3. Si la fecha es posterior al último valor de la tabla, solo se arrastra
       sobre sábados y domingos. Un día hábil sin valor al final de la tabla
       significa que el SII aún no lo publica: la solicitud queda en cuarentena
       y se resuelve en la próxima ejecución, con la tabla actualizada.
    """
    previous = [d for d in series if d <= day]
    if not previous:
        raise RateUnavailable(f"no hay dólar observado publicado antes del {day:%d-%m-%Y}")
    used = max(previous)
    age = (day - used).days
    if age > max_days:
        raise RateUnavailable(f"el último dólar publicado ({used:%d-%m-%Y}) tiene {age} días; máximo {max_days}")
    if day > max(series):
        pending = [used + timedelta(days=n) for n in range(1, age + 1)]
        business_day = next((d for d in pending if d.weekday() < 5), None)
        if business_day:
            raise RateUnavailable(
                f"el SII aún no publica el dólar del {business_day:%d-%m-%Y} "
                f"(último publicado: {used:%d-%m-%Y})"
            )
    return used, series[used], age
