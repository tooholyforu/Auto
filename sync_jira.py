#!/usr/bin/env python3
"""Собирает аллокацию за месяц из Jira и списка сотрудников и обновляет базу.

Доступ к Jira берется из переменных окружения, чтобы токен не попадал в историю команд:
    JIRA_URL, JIRA_TOKEN (или JIRA_USER и JIRA_PASSWORD), JIRA_CA_BUNDLE (необязательно).

Пример:
    python sync_jira.py --month 2026-09 --roster config/employees.csv
Без доступа к Jira можно взять ранее сохраненные списания:
    python sync_jira.py --month 2026-09 --roster config/employees.csv --worklogs-csv out/worklogs-2026-09.csv
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from allocation_calc import DEFAULT_MIN_HOURS, build_month, load_overrides, load_roster, month_bounds
from allocation_import import ERROR, WARNING, load_project_map, write_outputs
from jira_client import JiraClient, JiraError, build_session, load_worklogs_csv, save_worklogs_csv


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--month", required=True, help="месяц в формате ГГГГ-ММ")
    p.add_argument("--roster", type=Path, default=Path("config/employees.csv"), help="сотрудники из 1С")
    p.add_argument("--overrides", type=Path, default=Path("config/overrides.csv"),
                   help="фиксированные доли (используется, если файл существует)")
    p.add_argument("--project-map", type=Path, default=Path("config/jira_project_map.csv"),
                   help="соответствие проектов Jira проектам аллокации (используется, если файл существует)")
    p.add_argument("--worklogs-csv", type=Path, help="взять списания из файла вместо Jira")
    p.add_argument("--save-worklogs", type=Path, help="сохранить полученные из Jira списания в CSV")
    p.add_argument("--db", type=Path, default=Path("out/allocation.sqlite"), help="база, месяц в ней заменяется")
    p.add_argument("--out", type=Path, default=Path("out/sync"), help="папка для CSV и отчета")
    p.add_argument("--min-hours", type=float, default=DEFAULT_MIN_HOURS,
                   help="меньше стольких часов за месяц доли считаются ненадежными")
    p.add_argument("--strict", action="store_true", help="завершиться с кодом 1, если есть ошибки")
    args = p.parse_args(argv)

    start, end = month_bounds(args.month)
    if not args.roster.exists():
        raise SystemExit(f"не найден список сотрудников {args.roster}")
    roster = load_roster(args.roster)
    overrides = load_overrides(args.overrides) if args.overrides.exists() else None
    project_map = load_project_map(args.project_map) if args.project_map.exists() else None

    if args.worklogs_csv:
        worklogs = load_worklogs_csv(args.worklogs_csv, start, end)
    else:
        url = os.environ.get("JIRA_URL")
        if not url:
            raise SystemExit("не задан JIRA_URL (например https://jira.example.com)")
        try:
            session = build_session(os.environ.get("JIRA_TOKEN"), os.environ.get("JIRA_USER"),
                                    os.environ.get("JIRA_PASSWORD"), os.environ.get("JIRA_CA_BUNDLE"))
            worklogs = JiraClient(url, session).worklogs(start, end)
        except JiraError as exc:
            raise SystemExit(str(exc)) from None
        if args.save_worklogs:
            save_worklogs_csv(worklogs, args.save_worklogs)

    res = build_month(args.month, roster, worklogs, overrides, project_map, args.min_hours)
    write_outputs(res, args.out, db=args.db)

    errors = sum(i.level == ERROR for i in res.issues)
    warnings = sum(i.level == WARNING for i in res.issues)
    print(f"{args.month}: сотрудников {len(res.employee_months)}, списаний из Jira {len(worklogs)}, "
          f"строк аллокации {len(res.allocations)}")
    print(f"Ошибок: {errors}, предупреждений: {warnings}. Отчет: {args.out / 'report.md'}, база: {args.db}")
    return 1 if args.strict and errors else 0


if __name__ == "__main__":
    sys.exit(main())
