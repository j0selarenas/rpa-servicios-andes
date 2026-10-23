"""Orquestación de una ejecución: extracción, transformación y carga.

Errores de negocio (una solicitud con datos malos o sin tasa) -> cuarentena,
la ejecución sigue. Errores de sistema (SII caído, página cambiada, Excel sin
columnas, landing alterado) -> la ejecución termina en ERROR_SISTEMA sin cargar
nada nuevo; lo ya publicado queda intacto.
"""
from __future__ import annotations

import json
import logging
import uuid
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from filelock import FileLock, Timeout

from . import landing, storage
from .sii import (ExtractionError, RateUnavailable, StructureError, fetch_page, parse_rates,
                  to_decimal, uf_for, usd_for)
from .transform import InputFileError, parse_date, read_excel, to_bronze, validate

# Errores de sistema conocidos: se registran con su mensaje. Cualquier otro
# error es inesperado y se registra con el traceback completo.
KNOWN_ERRORS = (ExtractionError, StructureError, landing.IntegrityError, InputFileError)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _relative(path: Path, root: Path) -> str:
    try:
        return Path(path).relative_to(root).as_posix()
    except ValueError:
        return str(path)


def _start_log(run_id: str, folder: Path) -> logging.Logger:
    folder.mkdir(parents=True, exist_ok=True)
    log = logging.getLogger(f"rpa_andes.run.{run_id}")
    log.setLevel(logging.INFO)
    handler = logging.FileHandler(folder / "ejecucion.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter(f"%(asctime)s | %(levelname)-7s | {run_id[:8]} | %(message)s"))
    log.addHandler(handler)
    return log


def _years_needed(rows: list[dict], day: date, max_days: int) -> tuple[set[int], set[int]]:
    """Años de UF y dólar a descargar según las fechas del Excel (normalmente solo el actual)."""
    dates = []
    for row in rows:
        try:
            dates.append(parse_date(row["Fecha_Solicitud"], day))
        except (ValueError, TypeError):
            pass  # una fecha inválida se informa después, en la validación
    uf_years = {day.year} | {d.year for d in dates}
    usd_years = uf_years | {(d - timedelta(days=max_days)).year for d in dates}
    return uf_years, usd_years


def capture(config: dict, day: date, excel_path: Path, log, fetch=fetch_page) -> tuple[dict, bool]:
    """Guarda el Excel y las páginas del SII en landing. Una vez por día y archivo."""
    data_dir = config["rutas"]["datos"]
    content = Path(excel_path).read_bytes()
    existing = landing.manifest_path(data_dir, day, landing.sha256(content))
    if existing.exists():
        log.info("Este Excel ya fue capturado hoy; se reutiliza el snapshot sin volver al SII")
        return landing.read_manifest(existing), True

    excel = landing.save_source(data_dir, "solicitudes", day, "solicitudes.xlsx", content)
    log.info("Landing: Excel guardado en %s", excel["ruta"])
    rows = read_excel(data_dir / excel["ruta"], config["entrada"]["hoja"])
    uf_years, usd_years = _years_needed(rows, day, config["negocio"]["dias_max_arrastre_usd"])

    pages = []
    for indicator, key, years in (("UF", "url_uf", uf_years), ("USD", "url_dolar", usd_years)):
        for year in sorted(years):
            url = config["sii"][key].format(anio=year)
            log.info("Extracción: abriendo %s", url)
            page = fetch(url, config["navegador"], log)
            name = f"{'uf' if indicator == 'UF' else 'dolar'}{year}.html"
            record = landing.save_source(data_dir, "sii", day, name, page["html"])
            # Se valida la estructura antes de publicar el manifiesto: una página
            # rara queda en landing para revisarla, pero nunca como snapshot válido.
            values = parse_rates(page["html"], indicator, year)
            log.info("Landing: %s %s guardado en %s (%s valores, último %s)",
                     indicator, year, record["ruta"], len(values), max(values).isoformat())
            pages.append({**record, "indicador": indicator, "anio": year, "url": page["url"],
                          "http": page["http"], "capturado_en": page["capturado_en"]})

    manifest = landing.write_manifest(data_dir, day, excel, pages)
    log.info("Landing: snapshot %s publicado", manifest["snapshot_id"])
    return manifest, False


def enrich(valid: list[dict], rates: dict, config: dict) -> tuple[list[dict], list[dict]]:
    """Silver -> gold. Una solicitud sin tasa vigente va a cuarentena de enriquecimiento."""
    rules = config["negocio"]
    uf_unit = Decimal(1).scaleb(-rules["decimales_uf"])
    usd_unit = Decimal(1).scaleb(-rules["decimales_usd"])
    uf_series = {d: v for d, (v, _) in rates["UF"].items()}
    usd_series = {d: v for d, (v, _) in rates["USD"].items()}
    gold, missing = [], []
    for request in valid:
        day = request["fecha_solicitud"]
        try:
            uf_date, uf = uf_for(day, uf_series)
            usd_date, usd, age = usd_for(day, usd_series, rules["dias_max_arrastre_usd"])
        except RateUnavailable as exc:
            missing.append({"fila": request["fila"], "id_solicitud": request["id_solicitud"],
                            "etapa": "enriquecimiento", "codigos": "TASA_NO_DISPONIBLE",
                            "motivos": f"Fecha_Solicitud: {exc}"})
            continue
        amount = request["monto_clp"]
        gold.append({
            "id_solicitud": request["id_solicitud"], "rut": request["rut"],
            "nombre": request["nombre"], "apellido": request["apellido"], "email": request["email"],
            "fecha_solicitud": day.isoformat(),
            "monto_clp": amount,
            "monto_uf": (amount / uf).quantize(uf_unit, ROUND_HALF_UP),
            "monto_usd": (amount / usd).quantize(usd_unit, ROUND_HALF_UP),
            "valor_uf": uf, "fecha_uf": uf_date.isoformat(),
            "valor_usd": usd, "fecha_usd": usd_date.isoformat(), "dias_arrastre_usd": age,
            "fila_origen": request["fila"],
            "archivo_uf": rates["UF"][uf_date][1], "archivo_usd": rates["USD"][usd_date][1],
        })
    return gold, missing


def execute(config: dict, command: str, day: date, excel_path: Path | None = None,
            snapshot: str | None = None, fetch=fetch_page) -> dict:
    """Corre una ejecución completa y devuelve el resumen (también queda en resumen.json)."""
    run_id = str(uuid.uuid4())
    root = config["raiz"]
    data_dir = config["rutas"]["datos"]
    log_dir = config["rutas"]["logs"] / run_id
    log = _start_log(run_id, log_dir)
    summary = {"run_id": run_id, "comando": command, "fecha_proceso": day.isoformat(),
               "inicio": _now(), "estado": "EN_CURSO"}
    log.info("Inicio: comando=%s fecha=%s", command, day.isoformat())

    try:
        data_dir.mkdir(parents=True, exist_ok=True)
        with FileLock(str(data_dir / ".lock"), timeout=0):
            if command == "ejecutar":
                manifest, reused = capture(config, day, excel_path, log, fetch)
            else:
                manifest, reused = landing.find_manifest(data_dir, day, snapshot), True
                log.info("Reproceso del snapshot %s desde landing, sin conexión al SII", manifest["snapshot_id"])
            landing.verify_sources(data_dir, manifest)
            summary.update(snapshot_id=manifest["snapshot_id"], landing_reutilizado=reused)

            # Bronce
            rows = read_excel(data_dir / manifest["excel"]["ruta"], config["entrada"]["hoja"])
            rates, rate_rows = {"UF": {}, "USD": {}}, []
            for page in manifest["sii"]:
                html = (data_dir / page["ruta"]).read_bytes()
                for rate_date, raw in parse_rates(html, page["indicador"], page["anio"]).items():
                    value = to_decimal(raw)
                    rates[page["indicador"]][rate_date] = (value, page["ruta"])
                    rate_rows.append({"indicador": page["indicador"], "fecha": rate_date, "valor": value,
                                      "valor_original": raw, "archivo": page["ruta"], "url": page["url"]})
            log.info("Bronce: %s filas del Excel, %s valores UF, %s valores dólar",
                     len(rows), len(rates["UF"]), len(rates["USD"]))

            # Silver
            valid, rejected = validate(rows, cutoff=day)
            log.info("Silver: %s válidas, %s rechazadas por validación", len(valid), len(rejected))

            # Gold
            gold, missing = enrich(valid, rates, config)
            log.info("Gold: %s solicitudes convertidas, %s sin tasa vigente", len(gold), len(missing))
            quarantine = sorted(rejected + missing, key=lambda r: r["fila"])
            for item in quarantine:
                log.warning("Excepción de negocio | fila %s | %s | %s | %s",
                            item["fila"], item["id_solicitud"], item["codigos"], item["motivos"])

            if len(rows) != len(gold) + len(quarantine):
                raise RuntimeError(f"Conciliación fallida: {len(rows)} leídas != "
                                   f"{len(gold)} gold + {len(quarantine)} cuarentena")

            now = _now()
            conn = storage.connect(config["rutas"]["base_datos"])
            try:
                manifest_file = landing.manifest_path(data_dir, date.fromisoformat(manifest["fecha_extraccion"]),
                                                      manifest["excel"]["sha256"])
                storage.load_snapshot(conn, manifest, _relative(manifest_file, data_dir), {
                    "bronze": to_bronze(rows), "rates": rate_rows, "silver": valid,
                    "quarantine": quarantine, "gold": gold,
                }, run_id, now)
                log.info("Carga: snapshot %s publicado en SQLite", manifest["snapshot_id"])

                out_dir = config["rutas"]["salidas"]
                storage.export_run(conn, manifest["snapshot_id"], out_dir / run_id / "capas.xlsx")
                try:
                    storage.export_gold_actual(conn, out_dir / "gold_actual.xlsx")
                except PermissionError:
                    log.warning("No se pudo actualizar gold_actual.xlsx (¿está abierto en Excel?). "
                                "Los datos sí quedaron en la base; se actualizará en la próxima ejecución")
            finally:
                conn.close()

            summary.update({
                "estado": "OK_CON_CUARENTENA" if quarantine else "OK",
                "registros_leidos": len(rows),
                "registros_validos": len(valid),
                "registros_cuarentena": len(quarantine),
                "registros_gold": len(gold),
                "cuarentena_por_codigo": dict(Counter(c for q in quarantine for c in q["codigos"].split(";"))),
                "tasas": {ind: {"valores": len(v), "desde": min(v).isoformat(), "hasta": max(v).isoformat()}
                          for ind, v in rates.items()},
                "salidas": _relative(out_dir / run_id, root),
            })
    except Timeout:
        summary.update(estado="ERROR_SISTEMA", error="Hay otra ejecución en curso (data/.lock)")
        log.error(summary["error"])
    except KNOWN_ERRORS as exc:
        summary.update(estado="ERROR_SISTEMA", error=f"{type(exc).__name__}: {exc}")
        log.error("Error de sistema: %s", exc)
    except Exception as exc:
        summary.update(estado="ERROR_SISTEMA", error=f"{type(exc).__name__}: {exc}")
        log.exception("Error de sistema inesperado: %s", exc)
    finally:
        summary["fin"] = _now()
        summary["log"] = _relative(log_dir / "ejecucion.log", root)
        log.info("Fin: estado=%s leídas=%s válidas=%s cuarentena=%s gold=%s",
                 summary["estado"], summary.get("registros_leidos", 0), summary.get("registros_validos", 0),
                 summary.get("registros_cuarentena", 0), summary.get("registros_gold", 0))
        (log_dir / "resumen.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        for handler in list(log.handlers):
            handler.close()
            log.removeHandler(handler)
    return summary
