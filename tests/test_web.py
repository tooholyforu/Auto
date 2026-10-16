import sqlite3
import tempfile
import unittest
from datetime import date
from pathlib import Path

import allocation_import as ai
import jira_client as jc
import reconcile_jira as rec
import web_app
import webstore as ws
from allocation_calc import Employee

NAMES = ["Иванов Иван", "Петрова Анна", "Сидоров Сергей"]
POSITIONS = {"Иванов Иван": "Директор", "Петрова Анна": "Бухгалтер", "Сидоров Сергей": "Разработчик"}


def build_db(path: Path, august=None, september=None, evil_name=None) -> None:
    august = august or {"Иванов Иван": {"P1": .5, "P2": .5}, "Петрова Анна": {"P1": 1.0}, "Сидоров Сергей": {"P2": 1.0}}
    september = september or {"Иванов Иван": {"P1": .6, "P2": .4}, "Петрова Анна": {"P1": 1.0}, "Сидоров Сергей": {"P2": .5, "P3": .5}}
    res = ai.Result(months=["2026-08", "2026-09"])
    res.project_positions = {"2026-08": ["P1", "P2"], "2026-09": ["P3", "P1", "P2"]}       # порядок шапки Excel
    row = 0
    for month, data in (("2026-08", august), ("2026-09", september)):
        for name, shares in data.items():
            row += 1
            shown = evil_name if evil_name and name == "Иванов Иван" else name
            res.employee_months.append({"month": month, "employee": shown, "position": POSITIONS[name], "hire_date": "", "tag": "",
                                        "total_share": sum(shares.values()), "is_forecast": 0, "source": "excel", "source_row": row})
            for project, share in shares.items():
                res.allocations.append({"month": month, "employee": shown, "project": project, "share": share,
                                        "is_forecast": 0, "source_row": row})
    ai.write_sqlite(res, path)


class StoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "a.sqlite"
        build_db(self.path)
        self.db = ws.open_db(self.path)
        self.addCleanup(self.db.close)

    def ids(self):
        v = ws.get_month(self.db, "2026-09")
        return {r.name: r.employee_id for r in v.rows}, {n: i for i, n in v.projects}

    def test_view_order_history_candidates(self):
        v = ws.get_month(self.db, "2026-09")
        self.assertEqual([n for _, n in v.projects], ["P3", "P1", "P2"])                 # как в шапке Excel
        self.assertEqual([r.name for r in v.rows], NAMES)
        self.assertEqual(v.rows[0].history[0][0], "2026-08")
        self.assertEqual(v.rows[0].history[0][1], {"P1": .5, "P2": .5})
        self.assertEqual(v.previous, "2026-08")
        self.assertEqual(v.months, ["2026-08", "2026-09"])
        self.assertEqual(ws.problems(v.rows), [])

    def test_save_logs_only_real_changes_and_updates_totals(self):
        emp, prj = self.ids()
        values = {(emp["Иванов Иван"], prj["P1"]): .6,           # не изменилось
                  (emp["Иванов Иван"], prj["P2"]): .3,           # изменилось 0.4 -> 0.3
                  (emp["Петрова Анна"], prj["P3"]): .25,         # новая
                  (999, prj["P1"]): .5}                          # чужой id игнорируется
        result = ws.save_month(self.db, "2026-09", "Тимур", values)
        self.assertEqual(result.changed, 2)
        self.assertEqual(sorted(p.split(":")[0] for p in result.problems), ["Иванов Иван", "Петрова Анна"])
        log = ws.recent_log(self.db, "2026-09")
        self.assertEqual({(e["employee"], e["project"], e["old_value"], e["new_value"], e["user"]) for e in log},
                         {("Иванов Иван", "P2", .4, .3, "Тимур"), ("Петрова Анна", "P3", 0.0, .25, "Тимур")})
        row = self.db.execute("SELECT total_share, source FROM employee_months WHERE month='2026-09' AND employee_id=?",
                              (emp["Иванов Иван"],)).fetchone()
        self.assertEqual((row["total_share"], row["source"]), (.9, "web"))
        untouched = self.db.execute("SELECT source FROM employee_months WHERE month='2026-09' AND employee_id=?",
                                    (emp["Сидоров Сергей"],)).fetchone()
        self.assertEqual(untouched["source"], "excel")

    def test_zero_removes_share_and_range_is_checked(self):
        emp, prj = self.ids()
        ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P2"]): 0.0})
        self.assertNotIn(prj["P2"], ws.get_month(self.db, "2026-09").rows[0].shares)
        with self.assertRaises(ws.ValidationError):
            ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P1"]): 1.5})

    def test_lock_requires_complete_rows_and_blocks_edits(self):
        emp, prj = self.ids()
        ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P2"]): .3})
        with self.assertRaises(ws.ValidationError) as ctx:
            ws.lock_month(self.db, "2026-09", "т")
        self.assertIn("Иванов Иван", ctx.exception.problems[0])
        ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P2"]): .4})
        ws.lock_month(self.db, "2026-09", "Костя")
        self.assertTrue(ws.get_month(self.db, "2026-09").locked)
        for action in (lambda: ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P2"]): .1}),
                       lambda: ws.add_project(self.db, "2026-09", "т", "P9"),
                       lambda: ws.remove_employee(self.db, "2026-09", "т", emp["Иванов Иван"]),
                       lambda: ws.copy_from_previous(self.db, "2026-09", "т")):
            with self.assertRaises(ws.MonthLocked):
                action()
        ws.unlock_month(self.db, "2026-09", "Костя")
        ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P2"]): .4})
        actions = [e["action"] for e in ws.recent_log(self.db, "2026-09")]
        self.assertIn("месяц закрыт", actions)
        self.assertIn("месяц открыт заново", actions)

    def test_copy_creates_month_with_same_columns_and_does_not_overwrite(self):
        self.assertEqual(ws.copy_from_previous(self.db, "2026-10", "т"), 3)
        v = ws.get_month(self.db, "2026-10")
        self.assertEqual([n for _, n in v.projects], ["P3", "P1", "P2"])
        self.assertEqual([round(r.total, 2) for r in v.rows], [1.0, 1.0, 1.0])
        emp = {r.name: r.employee_id for r in v.rows}
        prj = {n: i for i, n in v.projects}
        ws.save_month(self.db, "2026-10", "т", {(emp["Иванов Иван"], prj["P1"]): .9, (emp["Иванов Иван"], prj["P2"]): .1})
        self.assertEqual(ws.copy_from_previous(self.db, "2026-10", "т"), 0)            # все уже заполнены
        self.assertAlmostEqual(ws.get_month(self.db, "2026-10").rows[0].shares[prj["P1"]], .9)

    def test_employees_and_projects(self):
        ws.remove_employee(self.db, "2026-09", "т", self.ids()[0]["Петрова Анна"])
        v = ws.get_month(self.db, "2026-09")
        self.assertNotIn("Петрова Анна", [r.name for r in v.rows])
        self.assertIn("Петрова Анна", [n for _, n in v.candidates])
        ws.add_employee(self.db, "2026-09", "т", employee_id=dict((n, i) for i, n in v.candidates)["Петрова Анна"])
        new_id = ws.add_employee(self.db, "2026-09", "т", name="  Новый   Человек ")
        again = ws.get_month(self.db, "2026-09")
        self.assertEqual(again.rows[-1].name, "Новый Человек")
        self.assertEqual(again.rows[-1].employee_id, new_id)
        self.assertEqual(again.rows[-2].position, "Бухгалтер")                         # должность подтянулась
        with self.assertRaises(ValueError):
            ws.add_employee(self.db, "2026-09", "т", name="Новый Человек")
        ws.add_project(self.db, "2026-09", "т", "P9")
        self.assertEqual([n for _, n in ws.get_month(self.db, "2026-09").projects][-1], "P9")
        actions = [e["action"] for e in ws.recent_log(self.db)]
        for expected in ("убран сотрудник", "добавлен сотрудник", "добавлен проект"):
            self.assertIn(expected, actions)

    def test_import_cannot_destroy_web_work(self):
        res = ai.Result(months=["2026-09"])
        ws.save_month(self.db, "2026-09", "т", {})                                       # без правок журнал пуст
        self.assertFalse(ai.has_web_edits(self.path))
        emp, prj = self.ids()
        ws.save_month(self.db, "2026-09", "т", {(emp["Иванов Иван"], prj["P1"]): .55, (emp["Иванов Иван"], prj["P2"]): .45})
        self.assertTrue(ai.has_web_edits(self.path))
        with self.assertRaisesRegex(SystemExit, "веб-интерфейс"):
            ai.write_sqlite(res, self.path)
        ws.lock_month(self.db, "2026-09", "т")
        with self.assertRaisesRegex(SystemExit, "закрыты"):
            ai.save_to_db(res, self.path)

    def test_bad_month_format(self):
        with self.assertRaises(ValueError):
            ws.get_month(self.db, "2026-13")


class AppTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)
        build_db(self.dir / "a.sqlite", evil_name="<script>alert(1)</script> Иван")
        self.app = web_app.create_app(web_app.Settings(self.dir / "a.sqlite"), secret="test")
        self.client = self.app.test_client()

    def form(self, **extra):
        db = ws.open_db(self.dir / "a.sqlite")
        v = ws.get_month(db, "2026-09")
        db.close()
        data = {"user": "Тимур"}
        for r in v.rows:
            for pid, _ in v.projects:
                data[f"v-{r.employee_id}-{pid}"] = web_app.fmt_pct(r.shares.get(pid))
        data.update(extra)
        return data, v

    def test_page_escapes_names_and_lists_everything(self):
        html = self.client.get("/m/2026-09").get_data(as_text=True)
        self.assertNotIn("<script>alert(1)</script>", html)
        self.assertIn("&lt;script&gt;alert(1)&lt;/script&gt;", html)
        self.assertEqual(html.count("<tr data-name"), 3)
        self.assertLess(html.index("<th>P3</th>"), html.index("<th>P1</th>"))

    def test_routes(self):
        self.assertEqual(self.client.get("/").headers["Location"], "/m/2026-09")
        self.assertEqual(self.client.get("/m/nonsense").status_code, 404)
        self.assertEqual(self.client.get("/go?month=2026-10").headers["Location"], "/m/2026-10")
        self.assertIn("пока нет данных", self.client.get("/m/2026-10").get_data(as_text=True))
        self.assertEqual(self.client.get("/log").status_code, 200)

    def test_save_roundtrip_and_percent_formats(self):
        data, v = self.form()
        emp, prj = {r.name.split()[-1]: r.employee_id for r in v.rows}, {n: i for i, n in v.projects}
        data[f"v-{emp['Иван']}-{prj['P1']}"] = "12,5"                       # запятая и дробное
        data[f"v-{emp['Иван']}-{prj['P2']}"] = "87.5%"
        response = self.client.post("/m/2026-09/save", data=data, follow_redirects=True).get_data(as_text=True)
        self.assertIn("изменено значений: 2", response)
        db = ws.open_db(self.dir / "a.sqlite")
        shares = ws.get_month(db, "2026-09").rows[0].shares
        db.close()
        self.assertEqual((shares[prj["P1"]], shares[prj["P2"]]), (.125, .875))

    def test_invalid_value_is_rejected_without_saving(self):
        data, v = self.form()
        key_ = next(k for k in data if k.startswith("v-"))
        data[key_] = "abc"
        response = self.client.post("/m/2026-09/save", data=data, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Не сохранено", response)
        db = ws.open_db(self.dir / "a.sqlite")
        self.assertEqual(ws.recent_log(db), [])
        db.close()

    def test_lock_flow_and_flash_messages(self):
        data, v = self.form()
        first, project = v.rows[0], v.projects[1][0]
        data[f"v-{first.employee_id}-{project}"] = "10"                      # сумма 70%
        self.client.post("/m/2026-09/save", data=data)
        text = self.client.post("/m/2026-09/lock", data={"user": "т"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Нельзя", text)
        data[f"v-{first.employee_id}-{project}"] = "60"                      # возвращаем 60 + 40 = 100%
        self.client.post("/m/2026-09/save", data=data)
        text = self.client.post("/m/2026-09/lock", data={"user": "т"}, follow_redirects=True).get_data(as_text=True)
        self.assertIn("Месяц закрыт", text)
        self.assertIn("disabled", text)
        text = self.client.post("/m/2026-09/save", data=data, follow_redirects=True).get_data(as_text=True)
        self.assertIn("закрыт", text)

    def test_cross_site_post_is_forbidden(self):
        r = self.client.post("/m/2026-09/lock", data={"user": "т"}, headers={"Origin": "https://evil.example"})
        self.assertEqual(r.status_code, 403)
        r = self.client.post("/m/2026-09/lock", data={"user": "т"}, headers={"Origin": "http://localhost"})
        self.assertNotEqual(r.status_code, 403)

    def test_jira_hints(self):
        (self.dir / "roster.csv").write_text("employee_id;name;jira_login\n1;Сидоров Сергей;ssid\n2;Петрова Анна;\n", encoding="utf-8-sig")
        (self.dir / "map.csv").write_text("raw_name;project;weight\nA;P1;1\nB;P2;1\n", encoding="utf-8-sig")
        (self.dir / "jira").mkdir()
        jc.save_assignments_csv([jc.Assignment("A-1", "A", "ssid", "ssid", ""), jc.Assignment("A-2", "A", "ssid", "ssid", "")],
                                self.dir / "jira" / "assignments-2026-09.csv")
        app = web_app.create_app(web_app.Settings(self.dir / "a.sqlite", self.dir / "roster.csv", self.dir / "map.csv", self.dir / "jira"), "t")
        html = app.test_client().get("/m/2026-09").get_data(as_text=True)
        self.assertIn("Данные Jira за 2026-09: задач 2", html)
        self.assertIn("2 зад.", html)                                       # у Сидорова P1: две задачи... есть
        self.assertIn('data-jira="0"', html)                                # у Сидорова P2 задач нет
        self.assertEqual(html.count('data-jira="'), 2)                      # у Петровой нет логина: не проверяется, подсказок нет
        self.assertIn("Данных Jira за 2026-10 нет", app.test_client().get("/m/2026-10").get_data(as_text=True))


class HelperTests(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(web_app.fmt_pct(0.3), "30")
        self.assertEqual(web_app.fmt_pct(0.125), "12,5")
        self.assertEqual(web_app.fmt_pct(None), "")
        self.assertEqual(web_app.parse_percent(" 12,5 "), 12.5)
        self.assertEqual(web_app.parse_percent(""), 0.0)
        for bad in ("abc", "101", "-1", "nan", "inf"):
            with self.assertRaises(ValueError):
                web_app.parse_percent(bad)
        self.assertEqual(web_app.next_month_code("2026-12"), "2027-01")
        self.assertEqual(web_app.next_month_code("2026-09"), "2026-10")

    def test_activity_by_project(self):
        roster = [Employee("1", "А", "", "a@x.ru", "alogin", None, None), Employee("2", "Б", "", "", "", None, None)]
        logs = [jc.Worklog("1", "A-1", "A", "alogin", "", "", date(2026, 9, 3), 7200)]
        tasks = [jc.Assignment("B-1", "B", "", "", "a@x.ru"), jc.Assignment("B-2", "B", "stranger", "", "")]
        activity, linked = rec.activity_by_project(roster, logs, tasks, {"a": [("P1", 1.0)], "b": [("P2", 1.0), ("P3", 0.5)]})
        self.assertEqual(activity, {"1": {"P1": {"tasks": 0, "hours": 2.0}, "P2": {"tasks": 1, "hours": 0.0}, "P3": {"tasks": 1, "hours": 0.0}}})
        self.assertEqual(linked, {"p1", "p2", "p3"})


if __name__ == "__main__":
    unittest.main()
