import csv
import tempfile
import unittest
from datetime import date
from pathlib import Path

import allocation_import as ai
import jira_client as jc
import reconcile_jira as rec
from test_sync import FakeJira, emp, issue, wl

MAP = {"dbo": [("ДБО/PWA", 1.0)], "abb": [("ABBID", 1.0)], "fact": [("Факторинг", 1.0)]}


def alloc(name, project, share, external_id=None):
    return {"employee": name, "external_id": external_id, "project": project, "share": share}


def assigned(login, project):
    return jc.Assignment(f"{project}-1", project, login, login, "")


class ReconcileTests(unittest.TestCase):
    def run_rec(self, allocations, worklogs=(), assignments=(), roster=None, **kw):
        roster = roster or [emp(1, "Сидоров Сергей", "ssidorov")]
        return rec.reconcile("2026-09", allocations, roster, list(worklogs), list(assignments), MAP, **kw)

    def test_allocation_without_jira_activity_is_flagged_with_context(self):
        out = self.run_rec([alloc("Сидоров Сергей", "ДБО/PWA", 0.6), alloc("Сидоров Сергей", "ABBID", 0.1)],
                           [wl("ssidorov", "DBO", 40)])
        self.assertEqual((out.checked, out.ok), (2, 1))
        case = out.flagged[0]
        self.assertEqual((case.employee, case.project, case.share, case.jira_projects), ("Сидоров Сергей", "ABBID", 0.1, "ABB"))
        self.assertIn("DBO: 40 ч", case.other_activity)

    def test_worklog_or_assignment_counts_as_activity(self):
        rows = [alloc("Сидоров Сергей", "ДБО/PWA", 0.5), alloc("Сидоров Сергей", "ABBID", 0.5)]
        out = self.run_rec(rows, [wl("ssidorov", "DBO", 1)], [assigned("SSIDOROV", "ABB")])
        self.assertEqual(out.flagged, [])

    def test_no_activity_at_all(self):
        out = self.run_rec([alloc("Сидоров Сергей", "Факторинг", 0.3)])
        self.assertEqual(out.flagged[0].other_activity, "в Jira за месяц вообще нет активности")

    def test_unverifiable_rows_are_reported_not_flagged(self):
        roster = [emp(1, "Сидоров Сергей", "ssidorov"), emp(2, "Без Логина", "")]
        out = self.run_rec([alloc("Сидоров Сергей", "Гарантии", 0.2), alloc("Без Логина", "ABBID", 1.0),
                            alloc("Неизвестный Человек", "ABBID", 1.0)], roster=roster)
        self.assertEqual(out.flagged, [])
        self.assertEqual(out.no_project, {"Гарантии": 1})
        self.assertEqual({n for n, _ in out.no_employee}, {"Без Логина", "Неизвестный Человек"})

    def test_zero_and_small_shares_are_skipped(self):
        rows = [alloc("Сидоров Сергей", "ABBID", 0.0), alloc("Сидоров Сергей", "Факторинг", 0.04)]
        self.assertEqual(self.run_rec(rows, min_share=0.05).checked, 0)
        self.assertEqual(len(self.run_rec(rows).flagged), 1)

    def test_employee_found_by_external_id_then_name(self):
        roster = [emp(7, "Полное Имя Отчество", "u7")]
        out = self.run_rec([alloc("Короткое Имя", "ABBID", 1.0, external_id="7")], roster=roster)
        self.assertEqual(len(out.flagged), 1)
        self.assertEqual(out.no_employee, [])

    def test_flagged_sorted_by_share(self):
        roster = [emp(1, "А", "a"), emp(2, "Б", "b")]
        out = self.run_rec([alloc("А", "ABBID", 0.1), alloc("Б", "ABBID", 0.7)], roster=roster)
        self.assertEqual([c.employee for c in out.flagged], ["Б", "А"])


class AssignmentsTests(unittest.TestCase):
    def test_client_returns_assignees_and_uses_updated_filter(self):
        fake = FakeJira([{"key": "DBO-5", "fields": {"project": {"key": "DBO"}, "assignee": {"name": "ssidorov", "key": "ssidorov"}}},
                         {"key": "ABB-1", "fields": {"project": {"key": "ABB"}, "assignee": {"name": "oorlov", "emailAddress": "o@x.ru"}}}])
        items = jc.JiraClient("https://j", fake, page_size=1).assignments(date(2026, 9, 1), date(2026, 10, 1))
        self.assertEqual([(a.project_key, a.login, a.email) for a in items], [("DBO", "ssidorov", ""), ("ABB", "oorlov", "o@x.ru")])
        self.assertIn('updated >= "2026-09-01"', fake.calls[0][1]["jql"])
        self.assertIn("assignee is not EMPTY", fake.calls[0][1]["jql"])


class CliTests(unittest.TestCase):
    def test_offline_run_writes_review_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            res = ai.Result(months=["2026-09"])
            res.employee_months += [{"month": "2026-09", "employee": "Сидоров Сергей", "external_id": "3", "position": "",
                                     "hire_date": "", "tag": "", "total_share": 1.0, "is_forecast": 0, "source": "excel",
                                     "source_row": 5}]
            res.allocations += [{"month": "2026-09", "employee": "Сидоров Сергей", "project": p, "share": s,
                                 "is_forecast": 0, "source_row": 5} for p, s in (("ДБО/PWA", 0.9), ("ABBID", 0.1))]
            ai.save_to_db(res, tmp / "a.sqlite")
            (tmp / "e.csv").write_text("employee_id;name;jira_login\n3;Сидоров Сергей Петрович;ssidorov\n", encoding="utf-8-sig")
            (tmp / "m.csv").write_text("raw_name;project;weight\nDBO;ДБО/PWA;1\nABB;ABBID;1\n", encoding="utf-8-sig")
            jc.save_worklogs_csv([jc.Worklog("1", "DBO-1", "DBO", "ssidorov", "ssidorov", "", date(2026, 9, 3), 36000)],
                                 tmp / "w.csv")
            code = rec.main(["--month", "2026-09", "--db", str(tmp / "a.sqlite"), "--roster", str(tmp / "e.csv"),
                             "--project-map", str(tmp / "m.csv"), "--worklogs-csv", str(tmp / "w.csv"),
                             "--out", str(tmp / "o")])
            self.assertEqual(code, 0)
            with open(tmp / "o" / "review-2026-09.csv", encoding="utf-8-sig", newline="") as f:
                rows = list(csv.DictReader(f, delimiter=";"))
            self.assertEqual([(r["employee"], r["project"], r["share"], r["decision"]) for r in rows],
                             [("Сидоров Сергей", "ABBID", "0,1", "")])
            self.assertIn("на ручной разбор: 1", (tmp / "o" / "report-2026-09.md").read_text(encoding="utf-8"))

    def test_missing_month_in_db_is_clear(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            ai.save_to_db(ai.Result(months=["2026-08"]), tmp / "a.sqlite")
            (tmp / "e.csv").write_text("employee_id;name;jira_login\n", encoding="utf-8-sig")
            (tmp / "m.csv").write_text("raw_name;project;weight\n", encoding="utf-8-sig")
            with self.assertRaisesRegex(SystemExit, "нет аллокации"):
                rec.main(["--month", "2026-09", "--db", str(tmp / "a.sqlite"), "--roster", str(tmp / "e.csv"),
                          "--project-map", str(tmp / "m.csv"), "--out", str(tmp)])


if __name__ == "__main__":
    unittest.main()
