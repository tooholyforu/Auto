import sqlite3
import tempfile
import unittest
from datetime import date
from pathlib import Path

import allocation_calc as calc
import allocation_import as ai
import jira_client as jc
import sync_jira

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"


class FakeResponse:
    def __init__(self, payload, status=200):
        self.payload, self.status_code = payload, status

    def json(self):
        return self.payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")


class FakeJira:
    """Имитирует поиск и списания Jira Server: постраничный ответ и обрезанные списания."""

    def __init__(self, issues, full_worklogs=None, status=200):
        self.issues, self.full, self.status, self.calls = issues, full_worklogs or {}, status, []

    def get(self, url, params=None, timeout=None):
        self.calls.append((url, dict(params or {})))
        if self.status != 200:
            return FakeResponse({}, self.status)
        if url.endswith("/rest/api/2/search"):
            start, size = params["startAt"], params["maxResults"]
            return FakeResponse({"issues": self.issues[start:start + size], "total": len(self.issues)})
        key = url.split("/issue/")[1].split("/")[0]
        entries = self.full[key]
        start, size = params["startAt"], params["maxResults"]
        return FakeResponse({"worklogs": entries[start:start + size], "total": len(entries)})


def entry(i, login, day, hours, email=""):
    return {"id": str(i), "started": f"2026-09-{day:02d}T10:00:00.000+0300", "timeSpentSeconds": int(hours * 3600),
            "author": {"name": login, "key": login, "emailAddress": email}}


def issue(key, entries, total=None):
    return {"key": key, "fields": {"project": {"key": key.split("-")[0]},
                                   "worklog": {"worklogs": entries, "total": len(entries) if total is None else total}}}


class JiraClientTests(unittest.TestCase):
    def fetch(self, fake, page_size=2):
        return jc.JiraClient("https://jira.example.com/", fake, page_size=page_size).worklogs(
            date(2026, 9, 1), date(2026, 10, 1))

    def test_pagination_and_filtering_by_month(self):
        fake = FakeJira([issue("A-1", [entry(1, "u1", 3, 2)]), issue("A-2", [entry(2, "u1", 4, 1)]),
                         issue("B-1", [entry(3, "u2", 5, 1), {**entry(4, "u2", 5, 1), "started": "2026-10-01T09:00:00.000+0300"}])])
        logs = self.fetch(fake)
        self.assertEqual(sorted(w.worklog_id for w in logs), ["1", "2", "3"])
        self.assertEqual([c[1]["startAt"] for c in fake.calls], [0, 2])
        jql = fake.calls[0][1]["jql"]
        self.assertIn('worklogDate >= "2026-09-01"', jql)
        self.assertIn('worklogDate < "2026-10-01"', jql)
        self.assertTrue(fake.calls[0][0].startswith("https://jira.example.com/rest/api/2/search"))

    def test_truncated_worklogs_are_fetched_in_full(self):
        full = [entry(i, "u1", 2, 1) for i in range(1, 6)]
        fake = FakeJira([issue("A-1", full[:2], total=5)], {"A-1": full})
        self.assertEqual(len(self.fetch(fake)), 5)

    def test_worklog_fields(self):
        fake = FakeJira([issue("DBO-7", [entry(1, "ssidorov", 3, 1.5, "s@example.com")])])
        w = self.fetch(fake)[0]
        self.assertEqual((w.project_key, w.issue_key, w.login, w.email, w.seconds, w.started),
                         ("DBO", "DBO-7", "ssidorov", "s@example.com", 5400, date(2026, 9, 3)))

    def test_access_denied_has_clear_message(self):
        with self.assertRaisesRegex(jc.JiraError, "токен и права"):
            self.fetch(FakeJira([], status=401))

    def test_session_requires_credentials(self):
        with self.assertRaises(jc.JiraError):
            jc.build_session()
        session = jc.build_session(token="t", ca_bundle="/etc/ca.pem")
        self.assertEqual(session.headers["Authorization"], "Bearer t")
        self.assertEqual(session.verify, "/etc/ca.pem")

    def test_csv_roundtrip_filters_by_period(self):
        logs = [jc.Worklog("1", "A-1", "A", "u", "u", "", date(2026, 9, 3), 3600),
                jc.Worklog("2", "A-1", "A", "u", "u", "", date(2026, 10, 3), 3600)]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "w.csv"
            jc.save_worklogs_csv(logs, path)
            back = jc.load_worklogs_csv(path, date(2026, 9, 1), date(2026, 10, 1))
        self.assertEqual(back, logs[:1])


def emp(i, name, login, hire=None, fire=None, email=""):
    return calc.Employee(str(i), name, "", email, login, hire, fire)


def wl(login, project, hours, day=3, email="", key=None):
    return jc.Worklog("x", f"{project}-1", project, login, key or login, email, date(2026, 9, day), int(hours * 3600))


class CalcTests(unittest.TestCase):
    def alloc(self, res, name):
        return {a["project"]: a["share"] for a in res.allocations if a["employee"] == name}

    def messages(self, res, level):
        return [i.message for i in res.issues if i.level == level]

    def test_shares_follow_hours_and_sum_to_one(self):
        res = calc.build_month("2026-09", [emp(1, "А А", "a")],
                               [wl("a", "P1", 1), wl("a", "P2", 1), wl("a", "P3", 1)], min_hours=1)
        shares = self.alloc(res, "А А")
        self.assertEqual(sum(shares.values()), 1.0)
        self.assertEqual(sorted(shares.values()), [0.3333, 0.3333, 0.3334])

    def test_apportion_always_sums_to_one(self):
        for weights in ({"a": 1, "b": 2, "c": 4}, {"a": 0.1, "b": 0.2, "c": 0.7, "d": 3.3}, {"a": 5}):
            self.assertAlmostEqual(sum(calc.apportion(weights).values()), 1.0, places=9)

    def test_match_by_login_key_or_email(self):
        roster = [emp(1, "А А", "a", email="a@x.ru"), emp(2, "Б Б", "b"), emp(3, "В В", "")]
        logs = [wl("A", "P1", 20), wl("zzz", "P1", 20, email="a@x.ru"), wl("", "P2", 20, key="b"),
                wl("", "P3", 20, email="nobody@x.ru")]
        res = calc.build_month("2026-09", roster, logs)
        self.assertEqual(self.alloc(res, "А А"), {"P1": 1.0})
        self.assertEqual(self.alloc(res, "Б Б"), {"P2": 1.0})
        self.assertTrue(any("nobody@x.ru" in m for m in self.messages(res, ai.WARNING)))

    def test_overrides_win_and_respect_period(self):
        roster = [emp(1, "А А", "a")]
        fixed = [calc.FixedShare("1", "P1", 0.4, "2026-09", None), calc.FixedShare("1", "P2", 0.6, "2026-09", "2026-09"),
                 calc.FixedShare("1", "OLD", 1.0, None, "2026-08")]
        res = calc.build_month("2026-09", roster, [wl("a", "P9", 30)], fixed)
        self.assertEqual(self.alloc(res, "А А"), {"P1": 0.4, "P2": 0.6})
        self.assertEqual(res.employee_months[0]["source"], "override")
        self.assertTrue(any("не использованы" in m for m in self.messages(res, ai.INFO)))
        self.assertEqual(calc.build_month("2026-10", roster, [], fixed).allocations[0]["project"], "P1")

    def test_override_sum_must_be_one(self):
        res = calc.build_month("2026-09", [emp(1, "А А", "a")], [], [calc.FixedShare("1", "P1", 0.5, None, None)])
        self.assertTrue(any("должно быть 1" in m for m in self.messages(res, ai.ERROR)))

    def test_employment_dates(self):
        roster = [emp(1, "Принят позже", "a", hire=date(2026, 10, 1)),
                  emp(2, "Принят в месяце", "b", hire=date(2026, 9, 30)),
                  emp(3, "Уволен до", "c", fire=date(2026, 8, 31)),
                  emp(4, "Уволен в месяце", "d", fire=date(2026, 9, 1))]
        logs = [wl(x, "P1", 20) for x in "abcd"]
        res = calc.build_month("2026-09", roster, logs)
        self.assertEqual({r["employee"] for r in res.employee_months}, {"Принят в месяце", "Уволен в месяце"})
        self.assertTrue(any("Уволен до" in m and "не учтены" in m for m in self.messages(res, ai.WARNING)))

    def test_no_data_and_low_hours_warnings(self):
        res = calc.build_month("2026-09", [emp(1, "Пустой", "a"), emp(2, "Мало", "b")], [wl("b", "P1", 3)])
        warnings = self.messages(res, ai.WARNING)
        self.assertTrue(any("Пустой" in m and "нет списаний" in m for m in warnings))
        self.assertTrue(any("Мало" in m and "3.0 ч" in m for m in warnings))
        self.assertEqual(next(r for r in res.employee_months if r["employee"] == "Пустой")["total_share"], 0)

    def test_project_map_merge_split_and_unmapped(self):
        mapping = {"p1": [("X", 1.0)], "p2": [("X", 1.0)], "shared": [("X", 0.5), ("Y", 0.5)]}
        res = calc.build_month("2026-09", [emp(1, "А А", "a")],
                               [wl("a", "P1", 10), wl("a", "P2", 10), wl("a", "SHARED", 20), wl("a", "NEW", 20)],
                               project_map=mapping)
        self.assertEqual(self.alloc(res, "А А"), {"NEW": 0.3333, "X": 0.5, "Y": 0.1667})
        self.assertEqual(len([m for m in self.messages(res, ai.ERROR) if "NEW" in m]), 1)

    def test_month_bounds(self):
        self.assertEqual(calc.month_bounds("2026-12"), (date(2026, 12, 1), date(2027, 1, 1)))
        with self.assertRaises(SystemExit):
            calc.month_bounds("2026-13")

    def test_roster_rejects_duplicates(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "e.csv"
            path.write_text("employee_id;name;jira_login\n1;А А;a\n2;Б Б;A\n", encoding="utf-8-sig")
            with self.assertRaisesRegex(SystemExit, "jira_login"):
                calc.load_roster(path)


class DatabaseTests(unittest.TestCase):
    def test_months_accumulate_and_rerun_replaces(self):
        roster = [emp(1, "А А", "a")]
        with tempfile.TemporaryDirectory() as tmp:
            db_path = Path(tmp) / "a.sqlite"
            ai.save_to_db(calc.build_month("2026-08", roster, [wl("a", "P1", 20)]), db_path)
            for hours in (20, 40):
                ai.save_to_db(calc.build_month("2026-09", roster, [wl("a", "P1", hours)]), db_path)
            db = sqlite3.connect(db_path)
            self.assertEqual(db.execute("select count(*) from employee_months").fetchone()[0], 2)
            self.assertEqual(db.execute("select hours from allocations where month='2026-09'").fetchall(), [(40.0,)])
            self.assertEqual(db.execute("select external_id from employees").fetchall(), [("1",)])
            self.assertEqual(db.execute("pragma foreign_key_check").fetchall(), [])
            db.close()


class CliTests(unittest.TestCase):
    def test_example_files_run_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmp:
            code = sync_jira.main([
                "--month", "2026-09", "--roster", str(EXAMPLES / "employees.csv"),
                "--overrides", str(EXAMPLES / "overrides.csv"), "--project-map", str(EXAMPLES / "jira_project_map.csv"),
                "--worklogs-csv", str(EXAMPLES / "worklogs.csv"), "--out", f"{tmp}/o", "--db", f"{tmp}/o/a.sqlite",
                "--strict"])
            self.assertEqual(code, 0)
            db = sqlite3.connect(f"{tmp}/o/a.sqlite")
            sums = db.execute("select round(sum(share), 4) from allocations group by employee_id").fetchall()
            self.assertTrue(all(s == (1.0,) for s in sums))
            self.assertEqual(len(sums), 5)       # без Волковой, у которой нет данных
            db.close()

    def test_jira_url_is_required_without_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaisesRegex(SystemExit, "JIRA_URL"):
                sync_jira.main(["--month", "2026-09", "--roster", str(EXAMPLES / "employees.csv"),
                                "--out", tmp, "--db", f"{tmp}/a.sqlite"])


if __name__ == "__main__":
    unittest.main()


class CheckTests(unittest.TestCase):
    def setUp(self):
        import check_jira
        self.check = check_jira

    def client(self, issues, status=200):
        fake = FakeJira(issues, status=status)
        original = fake.get

        def get(url, params=None, timeout=None):
            if url.endswith("/myself"):
                return FakeResponse({"name": "me"}, status)
            if url.endswith("/project"):
                return FakeResponse([{"key": "A"}, {"key": "B"}], status)
            return original(url, params, timeout)

        fake.get = get
        return jc.JiraClient("https://j", fake)

    def test_report_has_counts_but_no_logins_by_default(self):
        client = self.client([issue("A-1", [entry(1, "ssidorov", 3, 8), entry(2, "stranger", 4, 4)])])
        roster = [emp(1, "Сидоров", "ssidorov")]
        text = "\n".join(self.check.run_check(client, "2026-09", roster))
        self.assertIn("списаний за 2026-09: 2, часов 12, авторов 2", text)
        self.assertIn("найдено в списке сотрудников: 1 из 2", text)
        self.assertNotIn("stranger", text)
        self.assertIn("stranger", "\n".join(self.check.run_check(client, "2026-09", roster, verbose=True)))

    def test_no_worklogs_is_flagged(self):
        text = "\n".join(self.check.run_check(self.client([]), "2026-09"))
        self.assertIn("списаний за 2026-09 не найдено", text)

    def test_previous_month(self):
        self.assertEqual(self.check.previous_month(date(2026, 10, 9)), "2026-09")
        self.assertEqual(self.check.previous_month(date(2026, 1, 5)), "2025-12")

    def test_network_errors_are_explained(self):
        import requests
        self.assertIn("VPN", jc.explain(requests.exceptions.ConnectionError()))
        self.assertIn("JIRA_CA_BUNDLE", jc.explain(requests.exceptions.SSLError()))
