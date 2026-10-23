"""Ejecución completa en una carpeta temporal, con el SII simulado con las páginas reales guardadas."""
import json
import shutil
import sqlite3
import tempfile
import unittest
from contextlib import closing
from datetime import date, datetime, timezone
from pathlib import Path

from filelock import FileLock

from rpa_andes.config import load_config
from rpa_andes.pipeline import execute
from rpa_andes.sii import ExtractionError

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).parent / "fixtures"
DAY = date(2026, 10, 22)
TABLES = ("snapshots", "bronce_solicitudes", "bronce_tasas", "silver_solicitudes",
          "silver_tasas", "silver_cuarentena", "gold_solicitudes")


class FakeSii:
    """Reemplaza al navegador: entrega las páginas guardadas y cuenta las llamadas."""

    def __init__(self, usd_html=None):
        self.calls = 0
        self.usd_html = usd_html or (FIXTURES / "dolar2026.html").read_bytes()

    def __call__(self, url, browser, log):
        self.calls += 1
        html = (FIXTURES / "uf2026.html").read_bytes() if "/uf/" in url else self.usd_html
        return {"html": html, "url": url, "http": 200,
                "capturado_en": datetime.now(timezone.utc).isoformat(timespec="seconds")}


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder, ignore_errors=True)
        shutil.copy(ROOT / "config.toml", self.folder)
        shutil.copy(ROOT / "solicitudes.xlsx", self.folder)
        self.config = load_config(self.folder / "config.toml")
        self.excel = self.folder / "solicitudes.xlsx"

    def run_robot(self, command="ejecutar", day=DAY, fetch=None):
        return execute(self.config, command, day, self.excel, fetch=fetch or FakeSii())

    def db(self, sql, *params):
        with closing(sqlite3.connect(self.config["rutas"]["base_datos"])) as conn:
            return conn.execute(sql, params).fetchall()

    def counts(self):
        return {t: self.db(f"SELECT COUNT(*) FROM {t}")[0][0] for t in TABLES}

    def test_full_run_counts_and_quarantine(self):
        summary = self.run_robot()
        self.assertEqual(summary["estado"], "OK_CON_CUARENTENA")
        self.assertEqual((summary["registros_leidos"], summary["registros_validos"],
                          summary["registros_cuarentena"], summary["registros_gold"]), (30, 19, 13, 17))
        quarantined = {row[0]: row[1] for row in self.db("SELECT id_solicitud, codigos FROM silver_cuarentena")}
        self.assertEqual(quarantined["SOL-1007"], "RUT_INVALIDO")
        self.assertEqual(quarantined["SOL-1023"], "FECHA_INVALIDA")
        self.assertEqual(quarantined["SOL-1029"], "TASA_NO_DISPONIBLE")
        self.assertEqual(summary["cuarentena_por_codigo"]["ID_CONFLICTO"], 2)

    def test_conversions_and_weekend_rule(self):
        self.run_robot()
        gold = {r[0]: r[1:] for r in self.db(
            "SELECT id_solicitud, monto_uf, monto_usd, fecha_usd, dias_arrastre_usd FROM gold_solicitudes")}
        self.assertEqual(gold["SOL-1001"], (30.581, 1339.19, "2026-09-01", 0))
        self.assertEqual(gold["SOL-1005"], (17.6124, 771.32, "2026-09-04", 1))  # sábado
        self.assertEqual(gold["SOL-1016"], (10.008, 429.39, "2026-09-17", 2))  # feriado 19-09
        self.assertEqual(gold["SOL-1010"][0], 30.567)  # monto en texto 1.250.000,50

    def test_gold_traces_back_to_landing(self):
        self.run_robot()
        source, row, uf_file, sid = self.db(
            "SELECT archivo_origen, fila_origen, archivo_uf, snapshot_id FROM gold_solicitudes "
            "WHERE id_solicitud = 'SOL-1003'")[0]
        data = self.config["rutas"]["datos"]
        self.assertTrue((data / source).exists() and (data / uf_file).exists())
        manifest = self.db("SELECT manifiesto FROM snapshots WHERE snapshot_id = ?", sid)[0][0]
        self.assertEqual(json.loads((data / manifest).read_text("utf-8"))["excel"]["ruta"], source)
        self.assertEqual(self.db("SELECT id_solicitud FROM bronce_solicitudes WHERE snapshot_id = ? "
                                 "AND fila_origen = ?", sid, row)[0][0], "SOL-1003")

    def test_second_run_same_day_does_not_duplicate_or_call_sii(self):
        self.run_robot()
        before = self.counts()
        bronze_run = self.db("SELECT DISTINCT run_id FROM bronce_solicitudes")
        fake = FakeSii()
        summary = self.run_robot(fetch=fake)
        self.assertTrue(summary["landing_reutilizado"])
        self.assertEqual(fake.calls, 0)
        self.assertEqual(self.counts(), before)
        self.assertEqual(self.db("SELECT DISTINCT run_id FROM bronce_solicitudes"), bronze_run)

    def test_reprocess_rebuilds_from_landing_without_network(self):
        self.run_robot()
        before = self.db("SELECT * FROM gold_solicitudes ORDER BY id_solicitud")
        fake = FakeSii()
        summary = self.run_robot("reprocesar", fetch=fake)
        self.assertEqual(summary["estado"], "OK_CON_CUARENTENA")
        self.assertEqual(fake.calls, 0)
        after = self.db("SELECT * FROM gold_solicitudes ORDER BY id_solicitud")
        strip = lambda rows: [r[:-2] for r in rows]  # sin run_id ni procesado_en
        self.assertEqual(strip(after), strip(before))

    def test_modified_landing_file_blocks_reprocess(self):
        self.run_robot()
        before = self.counts()
        html = next((self.config["rutas"]["datos"] / "landing" / "sii").rglob("*dolar2026.html"))
        html.write_bytes(html.read_bytes().replace(b"972,60", b"999,99"))
        summary = self.run_robot("reprocesar")
        self.assertEqual(summary["estado"], "ERROR_SISTEMA")
        self.assertIn("fue modificado", summary["error"])
        self.assertEqual(self.counts(), before)

    def test_changed_sii_page_fails_without_loading(self):
        broken = (FIXTURES / "dolar2026.html").read_bytes().replace(b"<th>Oct</th>", b"<th>Octubre</th>")
        summary = self.run_robot(fetch=FakeSii(usd_html=broken))
        self.assertEqual(summary["estado"], "ERROR_SISTEMA")
        self.assertIn("StructureError", summary["error"])
        self.assertFalse(Path(self.config["rutas"]["base_datos"]).exists())
        # La página rara queda en landing para revisarla, pero no se publica snapshot.
        self.assertFalse(list((self.config["rutas"]["datos"] / "landing").rglob("manifiestos/*/*.json")))
        # El SII se arregla: la siguiente ejecución vuelve a intentar.
        fake = FakeSii()
        self.assertEqual(self.run_robot(fetch=fake)["estado"], "OK_CON_CUARENTENA")
        self.assertEqual(fake.calls, 2)

    def test_sii_down_is_a_system_error_with_log_and_summary(self):
        def down(url, browser, log):
            raise ExtractionError("El SII no respondió después de 3 intentos (timeout)")
        summary = self.run_robot(fetch=down)
        self.assertEqual(summary["estado"], "ERROR_SISTEMA")
        log_dir = self.config["rutas"]["logs"] / summary["run_id"]
        self.assertIn("Error de sistema", (log_dir / "ejecucion.log").read_text("utf-8"))
        self.assertEqual(json.loads((log_dir / "resumen.json").read_text("utf-8"))["estado"], "ERROR_SISTEMA")

    def test_concurrent_run_is_rejected(self):
        self.config["rutas"]["datos"].mkdir(parents=True)
        with FileLock(str(self.config["rutas"]["datos"] / ".lock")):
            summary = self.run_robot()
        self.assertEqual(summary["estado"], "ERROR_SISTEMA")
        self.assertIn("otra ejecución", summary["error"])

    def test_next_day_capture_resolves_pending_dollar(self):
        self.run_robot()
        # Al día siguiente el SII ya publicó el viernes 02-10.
        updated = (FIXTURES / "dolar2026.html").read_bytes().replace(
            b"<td style='text-align:right;'>936,86</td>\n   <td style='text-align:right;'>&nbsp;</td>",
            b"<td style='text-align:right;'>936,86</td>\n   <td style='text-align:right;'>975,10</td>", 1)
        self.assertNotEqual(updated, (FIXTURES / "dolar2026.html").read_bytes())
        summary = self.run_robot(day=date(2026, 10, 23), fetch=FakeSii(usd_html=updated))
        self.assertEqual(summary["registros_gold"], 19)
        current = dict(self.db("SELECT id_solicitud, fecha_usd FROM gold_actual"))
        self.assertEqual(len(current), 19)
        self.assertEqual(current["SOL-1029"], "2026-10-02")
        self.assertEqual(current["SOL-1030"], "2026-10-02")
        self.assertEqual(self.db("SELECT COUNT(*) FROM snapshots")[0][0], 2)

    def test_excel_outputs_are_written(self):
        summary = self.run_robot()
        self.assertTrue((self.folder / summary["salidas"] / "capas.xlsx").exists())
        self.assertTrue((self.config["rutas"]["salidas"] / "gold_actual.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
