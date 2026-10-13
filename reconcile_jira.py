#!/usr/bin/env python3
"""Сверяет заявленную аллокацию с активностью в Jira и выдает случаи для ручного разбора.

Правило: если в аллокации у человека за месяц указан проект (например 10%), а в Jira у него
по этому проекту нет задач (где он исполнитель и которые обновлялись в этом месяце), случай
попадает в список на ручной разбор. Ничего не исправляется само. Списания времени в Jira не
ведутся, поэтому по умолчанию они не используются (флаг --with-worklogs).

Аллокация берется из базы (allocation_import.py или sync_jira.py), сотрудники из списка 1С,
связь проектов с Jira из config/jira_project_map.csv (raw_name это ключ проекта Jira).

    python reconcile_jira.py --month 2026-09
Доступ к Jira: JIRA_URL, JIRA_TOKEN (или JIRA_USER и JIRA_PASSWORD), JIRA_CA_BUNDLE.
"""
from __future__ import annotations

import argparse
import contextlib
import os
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path

import requests

from allocation_calc import Employee, load_roster, month_bounds
from allocation_import import check_schema, key, load_project_map, num, write_csv
from jira_client import (Assignment, JiraClient, JiraError, Worklog, build_session, explain,
                         load_assignments_csv, load_worklogs_csv, save_assignments_csv)

CSV_HEADER = ["month", "employee", "project", "share", "jira_projects", "reason", "other_activity", "decision"]


@dataclass
class Case:
    month: str
    employee: str
    project: str
    share: float
    jira_projects: str
    reason: str
    other_activity: str


@dataclass
class Outcome:
    month: str
    checked: int = 0
    flagged: list[Case] = field(default_factory=list)
    no_employee: list[tuple[str, str]] = field(default_factory=list)      # (сотрудник, причина)
    no_project: dict[str, int] = field(default_factory=dict)              # проект -> число строк

    @property
    def ok(self) -> int:
        return self.checked - len(self.flagged)


def reconcile(month: str, allocations: list[dict], roster: list[Employee], worklogs: list[Worklog],
              assignments: list[Assignment], project_map: dict[str, list[tuple[str, float]]],
              min_share: float = 0.0) -> Outcome:
    by_id = {e.employee_id: e for e in roster}
    by_name = {key(e.name): e for e in roster}
    by_login = {key(e.jira_login): e for e in roster if e.jira_login}
    by_email = {key(e.email): e for e in roster if e.email}

    def owner(login: str, user_key: str, email: str) -> Employee | None:
        return by_login.get(key(login)) or by_login.get(key(user_key)) or by_email.get(key(email))

    # employee_id -> ключ проекта Jira -> {"hours", "tasks"}
    activity: dict[str, dict[str, dict[str, float]]] = {}

    def record(e: Employee | None, project_key: str, field_name: str, amount: float) -> None:
        if e is not None:
            slot = activity.setdefault(e.employee_id, {}).setdefault(key(project_key), {"hours": 0.0, "tasks": 0})
            slot[field_name] += amount

    for w in worklogs:
        record(owner(w.login, w.user_key, w.email), w.project_key, "hours", w.seconds / 3600)
    for a in assignments:
        record(owner(a.login, a.user_key, a.email), a.project_key, "tasks", 1)

    jira_keys: dict[str, set[str]] = {}
    for raw, targets in project_map.items():
        for project, _ in targets:
            jira_keys.setdefault(key(project), set()).add(raw)

    out = Outcome(month)
    for row in allocations:
        if row["share"] <= 0 or row["share"] < min_share:
            continue
        employee = by_id.get(row.get("external_id") or "") or by_name.get(key(row["employee"]))
        if employee is None:
            out.no_employee.append((row["employee"], "нет в списке сотрудников (проверьте написание ФИО)"))
            continue
        if not employee.jira_login:
            # пустой логин означает, что человек не работает в Jira (руководство, бухгалтерия), его не проверяем
            out.no_employee.append((row["employee"], "в списке сотрудников нет логина Jira, не проверяется"))
            continue
        keys = jira_keys.get(key(row["project"]))
        if not keys:
            out.no_project[row["project"]] = out.no_project.get(row["project"], 0) + 1
            continue

        out.checked += 1
        mine = activity.get(employee.employee_id, {})
        if any(k in mine for k in keys):
            continue
        others = sorted(((k, v) for k, v in mine.items()), key=lambda kv: -kv[1]["hours"])
        out.flagged.append(Case(
            month, row["employee"], row["project"], row["share"], ", ".join(sorted(k.upper() for k in keys)),
            "в Jira нет списаний и задач по этому проекту за месяц",
            "; ".join(f"{k.upper()}: {v['hours']:.0f} ч, задач {int(v['tasks'])}" for k, v in others)
            or "в Jira за месяц вообще нет активности"))
    out.flagged.sort(key=lambda c: (-c.share, c.employee))
    return out


def read_allocations(db_path: Path, month: str) -> list[dict]:
    with contextlib.closing(sqlite3.connect(db_path)) as db:
        check_schema(db, db_path)
        rows = db.execute(
            "SELECT e.name, e.external_id, p.name, a.share FROM allocations a "
            "JOIN employees e ON e.id = a.employee_id JOIN projects p ON p.id = a.project_id "
            "WHERE a.month = ? ORDER BY e.name, p.name", (month,)).fetchall()
    return [{"employee": n, "external_id": x, "project": p, "share": s} for n, x, p, s in rows]


def write_review(out: Outcome, directory: Path, activity_note: str) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    write_csv(directory / f"review-{out.month}.csv", CSV_HEADER,
              [[c.month, c.employee, c.project, num(c.share), c.jira_projects, c.reason, c.other_activity, ""]
               for c in out.flagged])
    lines = [f"# Сверка аллокации с Jira за {out.month}", "",
             f"- проверено строк аллокации: {out.checked}",
             f"- совпадает с активностью в Jira: {out.ok}",
             f"- на ручной разбор: {len(out.flagged)}",
             f"- не удалось проверить: {len(out.no_employee)} (сотрудник) и {sum(out.no_project.values())} (проект)",
             f"- что считалось активностью: {activity_note}", ""]
    if out.flagged:
        lines += ["## На ручной разбор", "", "| Сотрудник | Проект | Доля | Ключи Jira | Что есть в Jira |", "|---|---|---|---|---|"]
        lines += [f"| {c.employee} | {c.project} | {c.share:g} | {c.jira_projects} | {c.other_activity} |"
                  for c in out.flagged]
        lines.append("")
    if out.no_employee:
        lines += ["## Не проверены: нет данных о сотруднике", ""]
        lines += [f"- {name}: {why}" for name, why in sorted(set(out.no_employee))]
        lines.append("")
    if out.no_project:
        lines += ["## Не проверены: проект не связан с Jira", "",
                  "Добавьте проект в config/jira_project_map.csv, если он есть в Jira.", ""]
        lines += [f"- {p}: {n} стр." for p, n in sorted(out.no_project.items())]
        lines.append("")
    (directory / f"report-{out.month}.md").write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--month", required=True, help="месяц ГГГГ-ММ")
    p.add_argument("--db", type=Path, default=Path("out/allocation.sqlite"), help="база с аллокацией")
    p.add_argument("--roster", type=Path, default=Path("config/employees.csv"), help="сотрудники из 1С")
    p.add_argument("--project-map", type=Path, default=Path("config/jira_project_map.csv"),
                   help="соответствие проектов Jira проектам аллокации")
    p.add_argument("--assignments-csv", type=Path, help="взять задачи из файла вместо Jira")
    p.add_argument("--save-assignments", type=Path, help="сохранить полученные из Jira задачи в CSV")
    p.add_argument("--with-worklogs", action="store_true", help="учитывать и списания времени (если их ведут)")
    p.add_argument("--worklogs-csv", type=Path, help="взять списания из файла вместо Jira")
    p.add_argument("--min-share", type=float, default=0.0, help="не проверять доли меньше этой (например 0.05)")
    p.add_argument("--out", type=Path, default=Path("out/reconcile"), help="папка для результатов")
    args = p.parse_args(argv)

    start, end = month_bounds(args.month)
    for path, what in ((args.db, "база с аллокацией"), (args.roster, "список сотрудников"),
                       (args.project_map, "справочник проектов Jira")):
        if not path.exists():
            raise SystemExit(f"не найден файл: {what} ({path})")
    roster, project_map = load_roster(args.roster), load_project_map(args.project_map)
    allocations = read_allocations(args.db, args.month)
    if not allocations:
        raise SystemExit(f"в базе {args.db} нет аллокации за {args.month}")

    offline = args.assignments_csv or args.worklogs_csv
    if offline:
        assignments = load_assignments_csv(args.assignments_csv) if args.assignments_csv else []
        worklogs = load_worklogs_csv(args.worklogs_csv, start, end) if args.worklogs_csv else []
    else:
        url = os.environ.get("JIRA_URL")
        if not url:
            raise SystemExit("не задан JIRA_URL (например https://jira.example.com)")
        try:
            session = build_session(os.environ.get("JIRA_TOKEN"), os.environ.get("JIRA_USER"),
                                    os.environ.get("JIRA_PASSWORD"), os.environ.get("JIRA_CA_BUNDLE"))
            client = JiraClient(url, session)
            assignments = client.assignments(start, end)
            worklogs = client.worklogs(start, end) if args.with_worklogs else []
        except (JiraError, requests.RequestException) as exc:
            raise SystemExit(explain(exc)) from None
        if args.save_assignments:
            save_assignments_csv(assignments, args.save_assignments)
    note = "задачи, где человек исполнитель и которые обновлялись в этом месяце"
    if args.with_worklogs or args.worklogs_csv:
        note += ", и списания времени"

    outcome = reconcile(args.month, allocations, roster, worklogs, assignments, project_map, args.min_share)
    write_review(outcome, args.out, note)
    print(f"{args.month}: проверено {outcome.checked}, совпало {outcome.ok}, на ручной разбор {len(outcome.flagged)}, "
          f"не проверено {len(outcome.no_employee) + sum(outcome.no_project.values())}")
    print(f"Список: {args.out / f'review-{args.month}.csv'}, отчет: {args.out / f'report-{args.month}.md'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
