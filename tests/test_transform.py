import tempfile
import unittest
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from rpa_andes.transform import (InputFileError, normalize_email, normalize_rut, parse_amount,
                                 parse_date, read_excel, to_bronze, validate)

ROOT = Path(__file__).resolve().parents[1]
CUTOFF = date(2026, 10, 22)


def row(fila=2, **values):
    base = {"ID_Solicitud": "SOL-1", "RUT": "13.433.012-0", "Nombre": "Camila", "Apellido": "González",
            "Email": "camila@correo.cl", "Monto_CLP": 1000.0, "Fecha_Solicitud": datetime(2026, 9, 1)}
    return {"fila": fila, **base, **values}


class ValidatorTests(unittest.TestCase):
    def test_rut_accepts_common_formats(self):
        self.assertEqual(normalize_rut("13.433.012-0"), "13433012-0")
        self.assertEqual(normalize_rut("10530829-9"), "10530829-9")
        self.assertEqual(normalize_rut("146240399"), "14624039-9")
        self.assertEqual(normalize_rut("9.521.911-k"), "9521911-K")

    def test_rut_rejects_wrong_check_digit_and_bad_format(self):
        with self.assertRaisesRegex(ValueError, "dígito verificador"):
            normalize_rut("9.215.279-1")
        for bad in ("-13433012-0", "13,433,012-0", "abc", "1.23.456-7"):
            with self.assertRaises(ValueError, msg=bad):
                normalize_rut(bad)

    def test_email_is_trimmed_and_lowercased(self):
        self.assertEqual(normalize_email("  DIEGO.GUTIERREZ@Correo.CL "), "diego.gutierrez@correo.cl")
        for bad in ("tomas.silva.correo.cl", "a@b", "a..b@correo.cl", "a@@correo.cl"):
            with self.assertRaises(ValueError, msg=bad):
                normalize_email(bad)

    def test_amount_accepts_numbers_and_chilean_text(self):
        self.assertEqual(parse_amount(1250000.0), Decimal("1250000.00"))
        self.assertEqual(parse_amount("1.250.000,50"), Decimal("1250000.50"))
        self.assertEqual(parse_amount("$ 980.000"), Decimal("980000.00"))
        for bad in ("un millón", 0.0, -350000.0, "1,250.00", True):
            with self.assertRaises(ValueError, msg=bad):
                parse_amount(bad)

    def test_date_accepts_excel_and_text_formats(self):
        self.assertEqual(parse_date(datetime(2026, 9, 1), CUTOFF), date(2026, 9, 1))
        self.assertEqual(parse_date("14-09-2026", CUTOFF), date(2026, 9, 14))
        self.assertEqual(parse_date("2026/09/15", CUTOFF), date(2026, 9, 15))

    def test_date_rejects_impossible_unknown_and_future(self):
        with self.assertRaisesRegex(ValueError, "la fecha no existe"):
            parse_date("31-09-2026", CUTOFF)
        with self.assertRaisesRegex(ValueError, "formato de fecha"):
            parse_date("septiembre 2026", CUTOFF)
        with self.assertRaisesRegex(ValueError, "posterior"):
            parse_date("2026-10-23", CUTOFF)


class ValidateTests(unittest.TestCase):
    def test_all_reasons_of_a_row_are_reported(self):
        valid, rejected = validate([row(RUT=None, Email="x", Monto_CLP="mucho")], CUTOFF)
        self.assertEqual(valid, [])
        self.assertEqual(rejected[0]["codigos"], "CAMPO_VACIO;EMAIL_INVALIDO;MONTO_INVALIDO")
        self.assertIn("RUT: campo obligatorio vacío", rejected[0]["motivos"])

    def test_exact_duplicate_keeps_first_row(self):
        valid, rejected = validate([row(3), row(29)], CUTOFF)
        self.assertEqual([r["fila"] for r in valid], [3])
        self.assertEqual(rejected[0]["fila"], 29)
        self.assertEqual(rejected[0]["codigos"], "ID_DUPLICADO")

    def test_same_id_with_different_data_quarantines_both(self):
        valid, rejected = validate([row(7, Monto_CLP=1980000.0), row(30, Monto_CLP=2050000.0)], CUTOFF)
        self.assertEqual(valid, [])
        self.assertEqual([(r["fila"], r["codigos"]) for r in rejected], [(7, "ID_CONFLICTO"), (30, "ID_CONFLICTO")])

    def test_text_and_numeric_amounts_with_same_value_are_duplicates(self):
        valid, rejected = validate([row(2, Monto_CLP=980000.0), row(3, Monto_CLP="$ 980.000")], CUTOFF)
        self.assertEqual(len(valid), 1)
        self.assertEqual(rejected[0]["codigos"], "ID_DUPLICADO")


class ExcelTests(unittest.TestCase):
    def test_delivered_file_reads_30_rows_and_keeps_real_row_numbers(self):
        rows = read_excel(ROOT / "solicitudes.xlsx", "Solicitudes")
        self.assertEqual(len(rows), 30)  # 31 filas con contenido menos una completamente vacía
        by_id = {r["ID_Solicitud"]: r["fila"] for r in rows}
        self.assertEqual(by_id["SOL-1020"], 21)
        self.assertEqual(by_id["SOL-1021"], 23)  # la fila 22 está vacía

    def test_bronze_keeps_original_values_as_text(self):
        rows = read_excel(ROOT / "solicitudes.xlsx", "Solicitudes")
        bronze = {r["id_solicitud"]: r for r in to_bronze(rows)}
        self.assertEqual(bronze["SOL-1010"]["monto_clp"], "1.250.000,50")
        self.assertEqual(bronze["SOL-1001"]["monto_clp"], "1250000")
        self.assertEqual(bronze["SOL-1012"]["fecha_solicitud"], "14-09-2026")
        self.assertEqual(bronze["SOL-1024"]["email"], "  DIEGO.GUTIERREZ@Correo.CL ")

    def test_missing_column_stops_the_batch(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = Path(folder.name) / "incompleto.xlsx"
        workbook = Workbook()
        workbook.active.title = "Solicitudes"
        workbook.active.append(["ID_Solicitud", "RUT", "Nombre"])
        workbook.save(path)
        with self.assertRaisesRegex(InputFileError, "Faltan columnas"):
            read_excel(path, "Solicitudes")


if __name__ == "__main__":
    unittest.main()
