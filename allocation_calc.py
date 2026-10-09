"""Расчет аллокации за месяц из списка сотрудников (1С) и списаний времени (Jira).

Правила для каждого сотрудника, который работал в месяце:
1. Есть фиксированная аллокация в overrides.csv: берем ее (для тех, кто не списывает время).
2. Иначе доля проекта = часы в проекте / все часы сотрудника за месяц.
3. Иначе в отчете предупреждение: данных нет.
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from allocation_import import (ERROR, INFO, TOLERANCE, WARNING, Result, clean, key, parse_share,
                               read_csv)
from jira_client import Worklog

DEFAULT_MIN_HOURS = 16.0


@dataclass(frozen=True)
class Employee:
    employee_id: str
    name: str
    position: str
    email: str
    jira_login: str
    hire: date | None
    fire: date | None


@dataclass(frozen=True)
class FixedShare:
    employee_id: str
    project: str
    share: float
    first: str | None      # месяц YYYY-MM, с которого действует
    last: str | None       # месяц YYYY-MM, по который действует


def _date(text: str, where: str) -> date | None:
    text = clean(text or "")
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise SystemExit(f"{where}: не удалось разобрать дату {text!r}, нужен формат ГГГГ-ММ-ДД или ДД.ММ.ГГГГ")


def load_roster(path: Path) -> list[Employee]:
    """employees.csv: employee_id;name;jira_login (обязательные), position;email;hire_date;fire_date."""
    employees = []
    for r in read_csv(path, {"employee_id", "name", "jira_login"}):
        employees.append(Employee(clean(r["employee_id"]), clean(r["name"]), clean(r.get("position") or ""),
                                  clean(r.get("email") or ""), clean(r["jira_login"]),
                                  _date(r.get("hire_date"), path.name), _date(r.get("fire_date"), path.name)))
    for field, values in (("employee_id", [e.employee_id for e in employees]),
                          ("name", [key(e.name) for e in employees]),
                          ("jira_login", [key(e.jira_login) for e in employees if e.jira_login])):
        dupes = sorted({v for v in values if values.count(v) > 1})
        if dupes:
            raise SystemExit(f"{path}: повторяется {field}: {', '.join(dupes)}")
    return employees


def load_overrides(path: Path) -> list[FixedShare]:
    """overrides.csv: employee_id;project;share, необязательно from;to (месяцы ГГГГ-ММ)."""
    rows = []
    for r in read_csv(path, {"employee_id", "project", "share"}):
        try:
            share = parse_share(r["share"])
        except ValueError as exc:
            raise SystemExit(f"{path}: {r['employee_id']}, {r['project']}: {exc}") from None
        rows.append(FixedShare(clean(r["employee_id"]), clean(r["project"]), share or 0.0,
                               clean(r.get("from") or "") or None, clean(r.get("to") or "") or None))
    return rows


def month_bounds(month: str) -> tuple[date, date]:
    """YYYY-MM -> (первый день, первый день следующего месяца)."""
    try:
        year, number = map(int, month.split("-"))
        first = date(year, number, 1)
    except ValueError:
        raise SystemExit(f"месяц {month!r} должен быть в формате ГГГГ-ММ") from None
    return first, date(year + number // 12, number % 12 + 1, 1)


def apportion(weights: dict[str, float], digits: int = 4) -> dict[str, float]:
    """Делит 1 пропорционально весам так, чтобы округленные доли в сумме дали ровно 1."""
    total, unit = sum(weights.values()), 10 ** digits
    raw = {k: v / total * unit for k, v in weights.items()}
    base = {k: int(v) for k, v in raw.items()}
    for k in sorted(raw, key=lambda k: (-(raw[k] - base[k]), k))[:unit - sum(base.values())]:
        base[k] += 1
    return {k: v / unit for k, v in base.items()}


def build_month(month: str, roster: list[Employee], worklogs: list[Worklog],
                overrides: list[FixedShare] | None = None,
                project_map: dict[str, list[tuple[str, float]]] | None = None,
                min_hours: float = DEFAULT_MIN_HOURS) -> Result:
    first, after = month_bounds(month)
    res = Result(months=[month])
    overrides = overrides or []

    by_login = {key(e.jira_login): e for e in roster if e.jira_login}
    by_email = {key(e.email): e for e in roster if e.email}

    def active(e: Employee) -> bool:
        return (e.hire is None or e.hire < after) and (e.fire is None or e.fire >= first)

    # Часы по сотрудникам и проектам Jira
    hours: dict[str, dict[str, float]] = {}
    stray: dict[str, float] = {}
    for w in worklogs:
        owner = by_login.get(key(w.login)) or by_login.get(key(w.user_key)) or by_email.get(key(w.email))
        if owner is None:
            who = w.login or w.user_key or w.email or "автор не указан"
            stray[who] = stray.get(who, 0.0) + w.seconds / 3600
            continue
        per_project = hours.setdefault(owner.employee_id, {})
        per_project[w.project_key] = per_project.get(w.project_key, 0.0) + w.seconds / 3600
        res.raw_projects.setdefault(w.project_key, set()).add(month)

    for who, h in sorted(stray.items()):
        res.add(WARNING, month, 0, f"в Jira есть списания {h:.1f} ч от {who!r}, его нет в списке сотрудников")

    unmapped: set[str] = set()

    def to_projects(per_project: dict[str, float]) -> dict[str, float]:
        merged: dict[str, float] = {}
        for jira_key, h in per_project.items():
            if project_map is not None and key(jira_key) in project_map:
                targets = project_map[key(jira_key)]
            else:
                if project_map is not None and jira_key not in unmapped:
                    unmapped.add(jira_key)
                    res.add(ERROR, month, 0, f"проект Jira {jira_key!r} не описан в справочнике проектов")
                targets = [(jira_key, 1.0)]
            for project, weight in targets:
                merged[project] = merged.get(project, 0.0) + h * weight
        return merged

    for e in sorted(roster, key=lambda e: e.name):
        if not active(e):
            if e.employee_id in hours:
                res.add(WARNING, month, 0, f"{e.name}: есть списания в Jira, но по списку сотрудников "
                                           "в этом месяце не работает, часы не учтены")
            continue

        fixed = [o for o in overrides if o.employee_id == e.employee_id
                 and (o.first is None or o.first <= month) and (o.last is None or month <= o.last)]
        own_hours = hours.get(e.employee_id, {})
        total_hours = sum(own_hours.values())
        shares: dict[str, float] = {}
        hours_by_project: dict[str, float] = {}
        source = "jira"

        if fixed:
            source = "override"
            for o in fixed:
                shares[o.project] = shares.get(o.project, 0.0) + o.share
            if abs(sum(shares.values()) - 1) > TOLERANCE:
                res.add(ERROR, month, 0, f"{e.name}: фиксированные доли в сумме {sum(shares.values()):g}, должно быть 1")
            if total_hours:
                res.add(INFO, month, 0, f"{e.name}: заданы фиксированные доли, часы в Jira ({total_hours:.1f} ч) не использованы")
        elif total_hours:
            hours_by_project = to_projects(own_hours)
            shares = apportion(hours_by_project)
            if total_hours < min_hours:
                res.add(WARNING, month, 0, f"{e.name}: списано всего {total_hours:.1f} ч, доли могут быть неточными")
        else:
            res.add(WARNING, month, 0, f"{e.name}: нет списаний в Jira и не задана фиксированная аллокация")

        res.names.setdefault(key(e.name), e.name)
        res.employee_months.append({
            "month": month, "employee": e.name, "external_id": e.employee_id, "position": e.position,
            "hire_date": e.hire.isoformat() if e.hire else "", "tag": "",
            "total_share": round(sum(shares.values()), 6), "is_forecast": 0, "source": source, "source_row": 0})
        for project, share in sorted(shares.items()):
            res.allocations.append({
                "month": month, "employee": e.name, "project": project, "share": round(share, 6),
                "is_forecast": 0, "source_row": 0,
                "hours": round(hours_by_project[project], 2) if project in hours_by_project else None})
    return res
