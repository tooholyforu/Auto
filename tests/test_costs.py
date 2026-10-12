import sqlite3
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as idx

import allocation_import as ai
import cost_report as cr
import onec_import as oc


def sheet(path, rows):
    wb = openpyxl.Workbook()
    for r in rows:
        wb.active.append(r)
    wb.save(path)


def row(width=40, **cells):
    out = [None] * width
    for letter, value in cells.items():
        out[idx(letter) - 1] = value
    return out


def build(tmp: Path):
    """Вымышленные выгрузки 1С в той же раскладке, что настоящие."""
    sheet(tmp / "staff.xlsx", [
        row(A="Штатные сотрудники"), row(A="Организация", E="Тест ООО"), row(A="Подразделение"),
        row(A="№", B="Сотрудник", F="Табельный номер", M="Должность", AA="Оклад (тариф)", AB="Количество ставок",
            AC="Дата приема", AK="Территория.Районный коэффициент", AL="Процент северной надбавки"),
        row(A="Отдел А"),
        row(A=1, B="Иванов Иван Иванович", F="00001", M="Директор", AA=100000, AB=1, AC="01.01.2025"),
        row(A=2, B="Петрова Анна Сергеевна", F="00002", M="Бухгалтер", AA=50000, AB="2/5", AC="01.02.2025", AK=1.15),
        row(A="Отдел Б"),
        row(A=3, B="Сидоров Сергей Петрович", F="00003", M="Разработчик", AA=10000, AB=1, AC="01.09.2026", AK=1.5, AL=50),
    ])
    sheet(tmp / "payroll.xlsx", [
        row(A="Анализ зарплаты по сотрудникам (помесячно)"), row(A="Организация", D="Тест ООО"),
        row(A="Месяц", D="Сентябрь 2026"),
        row(A="Подразделение", T="Всего начислено"), row(A="Сотрудник", E="Дней"),
        row(A="Отдел А", T=300), row(A="Иванов Иван Иванович", T=100), row(A="Петрова Анна Сергеевна", T=200),
        row(A="Отдел Б", T=50), row(A="Сидоров Сергей Петрович", T=30), row(A="Уволенный Игорь Олегович", T=20),
        row(A="Итого", T=350),
    ])

    def taxes(name, period, values):
        rows = [row(A="Налоги и взносы (кратко)"), row(A="Период формирования отчета"), row(A=period),
                row(A="Организация", C="Начислено", F="НДФЛ", G="Единый тариф (до 2023 года - ОПС)", H="Несч. случаи, профзаболев."),
                row(A="Сотрудник"), row(A="Тест ООО", G=999, H=999)]
        rows += [row(A=n, G=g, H=h) for n, g, h in values]
        rows.append(row(A="Итого", G=sum(v[1] for v in values), H=sum(v[2] for v in values)))
        sheet(tmp / name, rows)

    taxes("taxes.xlsx", "Сентябрь 2026", [("Иванов Иван Иванович", 30, 1), ("Петрова Анна Сергеевна", 60, 2),
                                          ("Сидоров Сергей Петрович", 9, 0), ("Уволенный Игорь Олегович", 6, 0)])
    taxes("taxes_prev.xlsx", datetime(2026, 8, 1), [("Иванов Иван Иванович", 28, 1), ("Петрова Анна Сергеевна", 55, 2),
                                                   ("Уволенный Игорь Олегович", 5, 0)])
    sheet(tmp / "ndfl.xlsx", [
        row(A="Анализ НДФЛ по месяцам"), row(A="Организация", D="Доходы", G="Исчислено налога", H="Удержано налога"),
        row(A="Сотрудник"), row(A="Тест ООО", G=999), row(A="2026", G=999), row(A="Межрайонная ИФНС № 5", G=999),
        row(A="авг. 2026", G=7), row(A="Уволенный Игорь Олегович", G=7),
        row(A="сент. 2026", G=45), row(A="Иванов Иван Иванович", G=13), row(A="Петрова Анна Сергеевна", G=26),
        row(A="Сидоров Сергей Петрович", G=4), row(A="Итого", G=52),
    ])
    sheet(tmp / "card51.xlsx", [
        row(A="Карточка счета 51 за Сентябрь 2026 г."), row(A="Отбор: Статьи Равно Оклады"),
        row(A="Период", B="Документ", C="Аналитика Дт", D="Аналитика Кт", E="Дебет", H="Кредит"),
        row(E="Счет", H="Счет"), row(A="Сальдо на начало"),
        row(A="04.09.2026", B="Списание", C="Иванов Иван Иванович\n<...>", D="Оклады", E="70", H="51", I=80),
        row(A="04.09.2026", B="Списание", C="Петрова Анна Сергеевна\n<...>", D="Оклады", E="70", H="51", I=150),
        row(A="18.09.2026", B="Списание", C="Иванов Иван Иванович\n<...>", D="Оклады", E="70", H="51", I=10),
        row(A="Обороты за 18.09.26", E=0, H=240),
        row(A="Обороты за период и сальдо на конец", E=0, H=240, L=0),
    ])


class OnecTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        build(self.dir)

    def collect(self, **kw):
        d = self.dir
        return oc.collect_costs("2026-09", d / "staff.xlsx", d / "payroll.xlsx", d / "taxes.xlsx", d / "taxes_prev.xlsx",
                                d / "ndfl.xlsx", d / "card51.xlsx", **kw)

    def rec(self, res, name):
        return next(r for r in res.records if r.name.startswith(name))

    def test_amounts_per_employee(self):
        res = self.collect()
        a = self.rec(res, "Иванов")
        self.assertEqual((a.accrued, a.insurance_accrued, a.paid, a.ndfl, a.insurance_paid), (100, 31, 90, 13, 29))
        b = self.rec(res, "Петрова")
        self.assertEqual((b.accrued, b.insurance_accrued, b.paid, b.ndfl, b.insurance_paid), (200, 62, 150, 26, 57))

    def test_staff_rate_and_salary_rules(self):
        res = self.collect()
        self.assertEqual(self.rec(res, "Петрова").rate, 0.4)
        self.assertAlmostEqual(self.rec(res, "Петрова").base_salary, 57500)       # 50000 * 1.15
        self.assertAlmostEqual(self.rec(res, "Сидоров").base_salary, 20000)       # 10000*1.5 + 10000*50/100
        self.assertEqual(self.rec(res, "Иванов").base_salary, 100000)
        self.assertEqual(self.rec(res, "Петрова").department, "Отдел А")
        self.assertEqual(self.rec(res, "Сидоров").tab_no, "00003")

    def test_department_and_org_rows_are_not_people(self):
        names = {r.name for r in self.collect().records}
        self.assertEqual(names, {"Иванов Иван Иванович", "Петрова Анна Сергеевна", "Сидоров Сергей Петрович",
                                 "Уволенный Игорь Олегович"})

    def test_person_outside_staff_is_kept_with_note(self):
        res = self.collect()
        gone = self.rec(res, "Уволенный")
        self.assertFalse(gone.in_staff)
        self.assertEqual((gone.accrued, gone.insurance_paid), (20, 5))
        self.assertTrue(any("нет в штатном списке" in i.message and "Уволенный" in i.message for i in res.issues))

    def test_new_hire_without_previous_month_taxes_is_not_a_warning(self):
        warn = [i.message for i in self.collect().issues if i.level == ai.WARNING]
        # принят в сентябре, за август взносов быть не должно, а вот выплаты в карточке нет
        self.assertFalse(any("Сидоров" in m and "взносы за предыдущий месяц" in m for m in warn))

    def test_missing_payment_is_warning(self):
        warn = [i.message for i in self.collect().issues if i.level == ai.WARNING]
        self.assertTrue(any("Сидоров" in m and "карточка" in m for m in warn))

    def test_ndfl_sections_are_selected(self):
        self.assertEqual(self.rec(self.collect(), "Уволенный").ndfl, 0)
        both = self.collect(ndfl_months=["2026-08", "2026-09"])
        self.assertEqual(self.rec(both, "Уволенный").ndfl, 7)

    def test_totals_match_reports(self):
        warnings = [i.message for i in self.collect().issues if "не равна итогу" in i.message]
        self.assertEqual(warnings, [])

    def test_wrong_month_is_error(self):
        res = oc.collect_costs("2026-10", self.dir / "staff.xlsx", self.dir / "payroll.xlsx", self.dir / "taxes.xlsx",
                               self.dir / "taxes_prev.xlsx", self.dir / "ndfl.xlsx", self.dir / "card51.xlsx")
        errors = [i.message for i in res.issues if i.level == ai.ERROR]
        self.assertTrue(any("2026-09, а нужен 2026-10" in m for m in errors))

    def test_missing_column_is_clear(self):
        sheet(self.dir / "bad.xlsx", [row(A="Что-то", B="другое")])
        with self.assertRaisesRegex(SystemExit, "не найдены столбцы"):
            oc.read_payroll(self.dir / "bad.xlsx", set(), "2026-09", [])

    def test_cli_and_rerun_replaces_month(self):
        d = self.dir
        argv = ["--month", "2026-09", "--staff", str(d / "staff.xlsx"), "--payroll", str(d / "payroll.xlsx"),
                "--taxes", str(d / "taxes.xlsx"), "--taxes-prev", str(d / "taxes_prev.xlsx"), "--ndfl", str(d / "ndfl.xlsx"),
                "--card51", str(d / "card51.xlsx"), "--db", str(d / "c.sqlite"), "--out", str(d / "o")]
        self.assertEqual(oc.main(argv), 0)
        self.assertEqual(oc.main(argv), 0)
        db = sqlite3.connect(d / "c.sqlite")
        self.assertEqual(db.execute("select count(*) from employee_costs").fetchone()[0], 4)
        db.close()
        self.assertTrue((d / "o" / "costs-load-2026-09.md").exists())

    def test_helpers(self):
        self.assertEqual(oc.parse_rate("2/5"), 0.4)
        self.assertEqual(oc.parse_rate("0,5"), 0.5)
        self.assertEqual(oc.parse_rate(1), 1.0)
        self.assertIsNone(oc.parse_rate("много"))
        self.assertEqual(oc.prev_month("2026-01"), "2025-12")
        self.assertEqual(oc.month_of("сент. 2026"), "2026-09")
        self.assertEqual(oc.month_of(datetime(2026, 8, 1)), "2026-08")


def cost(name, rate=1.0, accrued=100.0, ins=30.0, paid=90.0, ndfl=13.0, ins_paid=28.0):
    return {"employee": name, "position": "", "department": "", "rate": rate, "base_salary": None, "accrued": accrued,
            "insurance_accrued": ins, "paid": paid, "ndfl": ndfl, "insurance_paid": ins_paid}


def alloc(name, project, share):
    return {"employee": name, "external_id": None, "project": project, "share": share}


class CostReportTests(unittest.TestCase):
    def test_costs_follow_shares_and_remainder_is_explicit(self):
        costs = [cost("А А"), cost("Б Б", accrued=200), cost("Нет Аллокации", accrued=50)]
        allocs = [alloc("А А", "P1", 0.6), alloc("А А", "P2", 0.4), alloc("Б Б", "P1", 0.5)]
        s = cr.compute("2026-09", allocs, costs)
        self.assertAlmostEqual(s.totals["P1"]["accrued"], 0.6 * 100 + 0.5 * 200)
        self.assertAlmostEqual(s.totals["P2"]["accrued"], 40)
        # Б Б распределен наполовину, Нет Аллокации совсем нет
        self.assertAlmostEqual(s.totals[cr.UNDISTRIBUTED]["accrued"], 0.5 * 200 + 50)
        total = sum(t["accrued"] for t in s.totals.values())
        self.assertAlmostEqual(total, 100 + 200 + 50)
        text = " ".join(i.message for i in s.issues)
        self.assertIn("Нет Аллокации", text)
        self.assertIn("сумма долей меньше 1: Б Б", text)

    def test_headcount_fte_and_outside(self):
        costs = [cost("А А", rate=0.5), cost("Б Б", rate=1.0)]
        allocs = [alloc("А А", "P1", 1.0), alloc("Б Б", "P1", 1.0), alloc("Внешний В", "P1", 1.0)]
        s = cr.compute("2026-09", allocs, costs, headcount_exclude={"Б Б"})
        p = s.totals["P1"]
        self.assertEqual((p["headcount"], p["fte"], p["outside"]), (1.0, 1.5, 1.0))       # исключение только из человек
        self.assertTrue(any("Внешний В" in i.message and "внешние" in i.message.lower() for i in s.issues))

    def test_aliases_and_oversubscribed_share(self):
        s = cr.compute("2026-09", [alloc("Короткое Имя", "P1", 0.8), alloc("Короткое Имя", "P2", 0.5)],
                       [cost("Полное Имя Отчество")], aliases={ai.key("Короткое Имя"): "Полное Имя Отчество"})
        self.assertAlmostEqual(s.totals["P1"]["accrued"] + s.totals["P2"]["accrued"], 130)
        self.assertTrue(any(i.level == ai.ERROR and "больше 1" in i.message for i in s.issues))

    def test_excel_summary_and_detail(self):
        s = cr.compute("2026-09", [alloc("А А", "P1", 1.0)], [cost("А А")])
        with tempfile.TemporaryDirectory() as tmp:
            plain, full = Path(tmp) / "a.xlsx", Path(tmp) / "b.xlsx"
            cr.write_excel(s, plain)
            cr.write_excel(s, full, detail=True)
            self.assertEqual(openpyxl.load_workbook(plain).sheetnames, ["По проектам"])
            wb = openpyxl.load_workbook(full)
            self.assertEqual(wb.sheetnames, ["По проектам", "Сотрудники"])
            ws = wb["По проектам"]
            labels = {r[0].value: r for r in ws.iter_rows(min_row=4)}
            self.assertEqual(labels["Оклады начислено"][1].value, 100)
            self.assertTrue(str(labels["Оклады начислено"][3].value).startswith("=SUM("))

    def test_cli_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            res = ai.Result(months=["2026-09"])
            res.employee_months.append({"month": "2026-09", "employee": "А А", "external_id": "1", "position": "", "hire_date": "",
                                        "tag": "", "total_share": 1.0, "is_forecast": 0, "source": "excel", "source_row": 1})
            res.allocations.append({"month": "2026-09", "employee": "А А", "project": "P1", "share": 1.0,
                                    "is_forecast": 0, "source_row": 1})
            ai.save_to_db(res, tmp / "a.sqlite")
            costs = oc.CostResult("2026-09", records=[oc.CostRecord("А А", "1", "", "", True, 1.0, None, 100, 30, 90, 13, 28)])
            oc.save_costs(costs, tmp / "c.sqlite")
            code = cr.main(["--month", "2026-09", "--alloc-db", str(tmp / "a.sqlite"), "--costs-db", str(tmp / "c.sqlite"),
                            "--aliases", str(tmp / "none"), "--headcount-exclude-file", str(tmp / "none2"), "--out", str(tmp / "o")])
            self.assertEqual(code, 0)
            self.assertTrue((tmp / "o" / "projects-2026-09.xlsx").exists())
            self.assertIn("Замечаний нет", (tmp / "o" / "projects-2026-09-notes.md").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
