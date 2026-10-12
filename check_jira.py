#!/usr/bin/env python3
"""Проверяет доступ к Jira и наличие задач с исполнителями. Запускать на машине с доступом к серверу (VPN).

Ничего не записывает. По умолчанию печатает только количества, без имен и логинов,
поэтому вывод можно показывать другим. Доступ берется из переменных окружения:
JIRA_URL, JIRA_TOKEN (или JIRA_USER и JIRA_PASSWORD), JIRA_CA_BUNDLE.

    python check_jira.py
    python check_jira.py --month 2026-09 --roster config/employees.csv
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

import requests

from allocation_calc import load_roster, month_bounds
from allocation_import import key
from jira_client import JiraClient, JiraError, build_session, explain


def previous_month(today: date) -> str:
    return f"{today.year - (today.month == 1)}-{(today.month - 2) % 12 + 1:02d}"


def run_check(client: JiraClient, month: str, roster=None, verbose: bool = False) -> list[str]:
    lines = []
    me = client.myself()
    lines.append(f"[ok] подключение есть, вы вошли как {me.get('name') or me.get('key') or '?'}")
    projects = client.projects()
    lines.append(f"[ok] доступных проектов: {len(projects)}" if projects else "[!] доступных проектов нет")
    if verbose:
        lines.append("     ключи: " + ", ".join(sorted(p["key"] for p in projects)))

    start, end = month_bounds(month)
    tasks = client.assignments(start, end)
    if not tasks:
        lines.append(f"[!] задач с исполнителем, обновленных за {month}, не найдено: проверьте права на проекты")
    else:
        people = {key(t.login or t.user_key or t.email) for t in tasks}
        lines.append(f"[ok] задач с исполнителем, обновленных за {month}: {len(tasks)}, исполнителей {len(people)}, "
                     f"проектов {len({t.project_key for t in tasks})}")
        if roster:
            logins = {key(e.jira_login) for e in roster if e.jira_login}
            emails = {key(e.email) for e in roster if e.email}
            matched = {key(t.login or t.user_key or t.email) for t in tasks
                       if key(t.login) in logins or key(t.user_key) in logins or key(t.email) in emails}
            lines.append(f"[{'ok' if matched == people else '!'}] исполнителей найдено в списке сотрудников: "
                         f"{len(matched)} из {len(people)}")
            if verbose and matched != people:
                lines.append("     не найдены: " + ", ".join(sorted(people - matched)))
    logs = client.worklogs(start, end)
    lines.append(f"[i] списаний времени за {month}: {len(logs)} (для сверки не нужны, время в Jira не ведут)")
    return lines


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--month", default=previous_month(date.today()), help="месяц ГГГГ-ММ (по умолчанию прошлый)")
    p.add_argument("--roster", type=Path, help="список сотрудников для сверки авторов")
    p.add_argument("--verbose", action="store_true", help="печатать ключи проектов и логины (личные данные)")
    args = p.parse_args(argv)

    url = os.environ.get("JIRA_URL")
    if not url:
        raise SystemExit("не задан JIRA_URL (например https://jira.example.com)")
    roster = load_roster(args.roster) if args.roster else None
    try:
        session = build_session(os.environ.get("JIRA_TOKEN"), os.environ.get("JIRA_USER"),
                                os.environ.get("JIRA_PASSWORD"), os.environ.get("JIRA_CA_BUNDLE"))
        print("\n".join(run_check(JiraClient(url, session), args.month, roster, args.verbose)))
    except (JiraError, requests.RequestException) as exc:
        print(f"[x] {explain(exc)}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
