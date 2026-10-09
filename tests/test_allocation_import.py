import sqlite3
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl

import allocation_import as ai


def build_workbook(path: Path) -> None:
    """Синтетический файл, повторяющий структуру настоящего (без реальных ФИО)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    rows = [
        # март: заголовок датой, шапка в двух строках
        [None, datetime(2026, 3, 26)],
        [None, "ФИО", "Должность", None, None, "аллокация по проектам"],
        [None, None, None, None, "Альфа", "Бета", "альфа 2"],
        [1, "Иванов Иван Иванович", "Директор", None, 0.6, 0.3, 0.1],
        [2, "Петров Петр Петрович", "Аналитик", None, 1, " ", None],
        ["дубль", "Иванов Иван Иванович", "Директор по продуктам", None, 0.6, 0.3, 0.1],
        [None, "Сидоров Сидор", "Тестировщик", None, 0.4, 0.4, None],       # сумма 0.8
        [None, "Кузнецов Кузьма Кузьмич", "Дизайнер"],                       # нет долей
        [None, None, None, None, "=SUM(E4:E8)", "=SUM(F4:F8)", "=SUM(G4:G8)"],
        # июнь: текст, запятая, формула, дубль без пометки, новый проект
        [None, "Июнь 2026"],
        [None, "ФИО", "Должность"],
        [None, None, None, "дата трудоустройства", "Альфа", "Бета", "Гамма+Дельта"],
        [1, "Иванов Иван Иванович", "Директор", datetime(2025, 1, 1), "0,5", 0.25, 0.25],
        [2, "Петров Петр Петрович", "Аналитик", None, "=1*0.5", 0.5, None],
        [3, "Петров  Петр Петрович ", "Аналитик", None, 1],                  # повтор без пометки
        [4, "Орлов Олег", "Инженер", None, "много"],                         # не число
        [None, "Итого", None, None, 2, 1, 0.25],
        # август: прогноз
        [None, "Август 2026 ПРОГНОЗ"],
        [None, "ФИО", "Должность"],
        [None, None, None, None, "Альфа", "Бета"],
        [1, "Иванов Иван", "Директор", None, 0.5, 0.5],
    ]
    for r in rows:
        ws.append(r)
    wb.save(path)


class ParseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "in.xlsx"
        build_workbook(self.path)
        self.res = ai.parse_workbook(self.path)

    def messages(self, level, month=None):
        return [i.message for i in self.res.issues if i.level == level and (month is None or i.month == month)]

    def alloc(self, month, employee):
        return {a["project"]: a["share"] for a in self.res.allocations
                if a["month"] == month and a["employee"] == employee}

    def test_months_and_forecast(self):
        self.assertEqual(self.res.months, ["2026-03", "2026-06", "2026-08"])
        forecast = {r["month"]: r["is_forecast"] for r in self.res.employee_months}
        self.assertEqual(forecast["2026-08"], 1)
        self.assertEqual(forecast["2026-03"], 0)

    def test_blank_and_zero_cells_are_skipped(self):
        self.assertEqual(self.alloc("2026-03", "Петров Петр Петрович"), {"Альфа": 1.0})

    def test_totals_rows_are_not_loaded(self):
        names = {r["employee"] for r in self.res.employee_months}
        self.assertNotIn("Итого", names)
        self.assertEqual(self.messages(ai.ERROR, "2026-03"),
                         ["Сидоров Сидор: сумма долей 0.8, должна быть 1"])

    def test_duplicate_with_tag_is_skipped(self):
        march = [r for r in self.res.employee_months if r["month"] == "2026-03"]
        self.assertEqual(sum(r["employee"] == "Иванов Иван Иванович" for r in march), 1)
        self.assertTrue(any("дубль строки 4" in m for m in self.messages(ai.INFO, "2026-03")))

    def test_duplicate_without_tag_is_error(self):
        errors = self.messages(ai.ERROR, "2026-06")
        self.assertTrue(any("повторяется" in m for m in errors))
        self.assertEqual(self.alloc("2026-06", "Петров Петр Петрович"), {"Альфа": 0.5, "Бета": 0.5})

    def test_comma_and_formula_values(self):
        self.assertEqual(self.alloc("2026-06", "Иванов Иван Иванович")["Альфа"], 0.5)
        self.assertEqual(self.alloc("2026-06", "Петров Петр Петрович")["Альфа"], 0.5)

    def test_text_in_share_cell_is_error_and_row_not_loaded(self):
        self.assertTrue(any("E16" in m and "не число" in m for m in self.messages(ai.ERROR, "2026-06")))
        self.assertNotIn("Орлов Олег", {r["employee"] for r in self.res.employee_months})

    def test_no_allocation_is_warning(self):
        self.assertTrue(any("Кузнецов" in m and "нет ни одной доли" in m for m in self.messages(ai.WARNING)))

    def test_hire_date_and_merged_project_hint(self):
        row = next(r for r in self.res.employee_months
                   if r["month"] == "2026-06" and r["employee"].startswith("Иванов Иван Иванович"))
        self.assertEqual(row["hire_date"], "2025-01-01")
        self.assertTrue(any("Гамма+Дельта" in m for m in self.messages(ai.INFO)))

    def test_alias_suggestion(self):
        self.assertTrue(any("'Иванов Иван'" in m and "возможно один человек" in m
                            for m in self.messages(ai.INFO)))

    def test_default_project_names_merge_case_variants(self):
        projects = {a["project"] for a in self.res.allocations}
        self.assertIn("Альфа", projects)
        self.assertNotIn("альфа", projects)


class MappingTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        self.path = self.dir / "in.xlsx"
        build_workbook(self.path)

    def test_project_map_split_rename_and_unmapped(self):
        mapping = {ai.key("Альфа"): [("Альфа", 1.0)],
                   ai.key("Бета"): [("Бета-1", 0.5), ("Бета-2", 0.5)],
                   ai.key("Гамма+Дельта"): [("Гамма", 0.5), ("Дельта", 0.5)]}
        res = ai.parse_workbook(self.path, project_map=mapping)
        alloc = {a["project"]: a["share"] for a in res.allocations
                 if a["month"] == "2026-06" and a["employee"].startswith("Иванов Иван Иванович")}
        self.assertEqual(alloc, {"Альфа": 0.5, "Бета-1": 0.125, "Бета-2": 0.125, "Гамма": 0.125, "Дельта": 0.125})
        unmapped = [i.message for i in res.issues if i.level == ai.ERROR and "не описан" in i.message]
        self.assertTrue(any("альфа 2" in m for m in unmapped))

    def test_aliases_merge_employee_spellings(self):
        res = ai.parse_workbook(self.path, aliases={ai.key("Иванов Иван"): "Иванов Иван Иванович"})
        august = [r["employee"] for r in res.employee_months if r["month"] == "2026-08"]
        self.assertEqual(august, ["Иванов Иван Иванович"])
        self.assertFalse(any("возможно один человек" in i.message and "'Иванов Иван'" in i.message
                             for i in res.issues))

    def test_bad_weights_stop_the_import(self):
        path = self.dir / "map.csv"
        path.write_text("raw_name;project;weight\nАльфа;А;0,5\nАльфа;Б;0,4\n", encoding="utf-8-sig")
        with self.assertRaises(SystemExit):
            ai.load_project_map(path)


class OutputTests(unittest.TestCase):
    def test_outputs_and_sqlite(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / "in.xlsx", Path(tmp) / "out"
            build_workbook(src)
            self.assertEqual(ai.main([str(src), "--out", str(out),
                                      "--project-map", str(Path(tmp) / "none.csv"),
                                      "--aliases", str(Path(tmp) / "none2.csv")]), 0)
            for name in ("allocations.csv", "employee_months.csv", "allocation.sqlite", "report.md",
                         "project_map.suggested.csv", "employee_aliases.suggested.csv"):
                self.assertTrue((out / name).exists(), name)
            db = sqlite3.connect(out / "allocation.sqlite")
            self.assertEqual(db.execute("pragma foreign_key_check").fetchall(), [])
            total = db.execute("select round(sum(share), 4) from allocations where month='2026-08'").fetchone()[0]
            self.assertEqual(total, 1.0)
            db.close()

    def test_strict_mode_fails_on_errors(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "in.xlsx"
            build_workbook(src)
            code = ai.main([str(src), "--out", str(Path(tmp) / "o"), "--strict",
                            "--project-map", str(Path(tmp) / "x"), "--aliases", str(Path(tmp) / "y")])
            self.assertEqual(code, 1)


class ValueTests(unittest.TestCase):
    def test_parse_share(self):
        self.assertIsNone(ai.parse_share(" "))
        self.assertEqual(ai.parse_share("50%"), 0.5)
        self.assertEqual(ai.parse_share(0.30000000000000004), 0.3)
        for bad in (-0.1, 1.5, "abc", True):
            with self.assertRaises(ValueError):
                ai.parse_share(bad)

    def test_parse_title(self):
        self.assertEqual(ai.parse_title("май.26"), (2026, 5, False))
        self.assertEqual(ai.parse_title("Октябрь 26"), (2026, 10, False))
        self.assertEqual(ai.parse_title("Август 2026 ПРОГНОЗ"), (2026, 8, True))
        self.assertEqual(ai.parse_title(datetime(2026, 4, 26)), (2026, 4, False))
        self.assertIsNone(ai.parse_title("Кратко:"))
        self.assertIsNone(ai.parse_title("Иванов Иван"))
        self.assertIsNone(ai.parse_title("итого"))


if __name__ == "__main__":
    unittest.main()
