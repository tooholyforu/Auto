"""Данные веб-интерфейса аллокации: чтение и запись месяца в базе allocation.sqlite.

Без веб-фреймворка, поэтому логика проверяется обычными тестами. Основные правила:
- у сотрудника за месяц сумма долей должна быть 1 (допуск 0,005), иначе месяц нельзя закрыть;
- закрытый месяц не меняется, пока его явно не откроют заново;
- каждое изменение пишется в журнал: кто, когда, что было, что стало.
"""
from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from allocation_import import SCHEMA, TOLERANCE, check_schema, clean

MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")

EXTRA_SCHEMA = """
CREATE TABLE IF NOT EXISTS month_status (
    month TEXT PRIMARY KEY,
    locked INTEGER NOT NULL DEFAULT 0,
    changed_by TEXT, changed_at TEXT);
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY,
    at TEXT NOT NULL, user TEXT NOT NULL, month TEXT NOT NULL, action TEXT NOT NULL,
    employee TEXT, project TEXT, old_value REAL, new_value REAL, note TEXT);
"""


class MonthLocked(Exception):
    pass


class ValidationError(Exception):
    def __init__(self, problems: list[str]):
        super().__init__("; ".join(problems))
        self.problems = problems


@dataclass
class Row:
    employee_id: int
    name: str
    position: str
    shares: dict[int, float]                                   # id проекта -> доля
    external_id: str | None = None                             # табельный номер из 1С
    history: list[tuple[str, dict[str, float]]] = field(default_factory=list)   # (месяц, {проект: доля}), свежие первыми

    @property
    def total(self) -> float:
        return round(sum(self.shares.values()), 6)

    @property
    def complete(self) -> bool:
        return abs(self.total - 1) <= TOLERANCE


@dataclass
class MonthView:
    month: str
    locked: bool
    projects: list[tuple[int, str]]
    rows: list[Row]
    months: list[str]
    previous: str | None
    candidates: list[tuple[int, str]]                          # сотрудники, которых в месяце нет


@dataclass
class SaveResult:
    changed: int
    problems: list[str]


def check_month(month: str) -> str:
    if not MONTH_RE.match(month or ""):
        raise ValueError(f"месяц {month!r} должен быть в формате ГГГГ-ММ")
    return month


def open_db(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path)
    db.row_factory = sqlite3.Row
    check_schema(db, path)
    db.executescript(SCHEMA)
    db.executescript(EXTRA_SCHEMA)
    return db


# --------------------------------------------------------------------------
# Чтение
# --------------------------------------------------------------------------

def list_months(db: sqlite3.Connection) -> list[str]:
    rows = db.execute("SELECT month FROM employee_months UNION SELECT month FROM month_status ORDER BY month")
    return [r[0] for r in rows]


def previous_month_with_data(db: sqlite3.Connection, month: str) -> str | None:
    return db.execute("SELECT MAX(month) FROM employee_months WHERE month < ?", (month,)).fetchone()[0]


def is_locked(db: sqlite3.Connection, month: str) -> bool:
    row = db.execute("SELECT locked FROM month_status WHERE month = ?", (month,)).fetchone()
    return bool(row and row["locked"])


def _project_ids(db: sqlite3.Connection, month: str, previous: str | None) -> list[tuple[int, str]]:
    """Столбцы месяца в порядке шапки исходного Excel; проекты, добавленные позже, идут в конце."""
    def used(m: str) -> set[int]:
        ids = {r[0] for r in db.execute("SELECT DISTINCT project_id FROM allocations WHERE month = ?", (m,))}
        return ids | {r[0] for r in db.execute("SELECT project_id FROM month_projects WHERE month = ?", (m,))}

    basis, ids = month, used(month)
    if not ids and previous:
        basis, ids = previous, used(previous)
    if not ids:
        return []
    marks = ",".join("?" * len(ids))
    query = (f"SELECT p.id, p.name FROM projects p LEFT JOIN month_projects mp ON mp.project_id = p.id AND mp.month = ? "
             f"WHERE p.id IN ({marks}) ORDER BY mp.position IS NULL, mp.position, p.id")
    return [(r["id"], r["name"]) for r in db.execute(query, (basis, *ids))]


def get_month(db: sqlite3.Connection, month: str, history_depth: int = 4) -> MonthView:
    check_month(month)
    previous = previous_month_with_data(db, month)
    people = db.execute(
        "SELECT m.employee_id, e.name, e.external_id, COALESCE(m.position, '') AS position FROM employee_months m "
        "JOIN employees e ON e.id = m.employee_id WHERE m.month = ? "
        "ORDER BY CASE WHEN m.source_row = 0 THEN 1 ELSE 0 END, m.source_row, m.rowid", (month,)).fetchall()
    shares: dict[int, dict[int, float]] = {}
    for r in db.execute("SELECT employee_id, project_id, share FROM allocations WHERE month = ?", (month,)):
        shares.setdefault(r["employee_id"], {})[r["project_id"]] = r["share"]

    history: dict[int, dict[str, dict[str, float]]] = {}
    if people:
        marks = ",".join("?" * len(people))
        query = ("SELECT a.employee_id, a.month, p.name, a.share FROM allocations a JOIN projects p ON p.id = a.project_id "
                 f"WHERE a.month < ? AND a.employee_id IN ({marks}) ORDER BY a.month DESC, a.share DESC")
        for r in db.execute(query, (month, *[p["employee_id"] for p in people])):
            history.setdefault(r["employee_id"], {}).setdefault(r["month"], {})[r["name"]] = r["share"]

    rows = [Row(p["employee_id"], p["name"], p["position"], shares.get(p["employee_id"], {}), p["external_id"],
                list(history.get(p["employee_id"], {}).items())[:history_depth]) for p in people]
    present = {r.employee_id for r in rows}
    candidates = [(r["id"], r["name"]) for r in db.execute("SELECT id, name FROM employees ORDER BY name") if r["id"] not in present]
    months = list_months(db)
    return MonthView(month, is_locked(db, month), _project_ids(db, month, previous), rows, months, previous, candidates)


def problems(rows: list[Row]) -> list[str]:
    """Сотрудники, у которых сумма долей не равна 1."""
    return [f"{r.name}: {r.total * 100:g}%" for r in rows if not r.complete]


def recent_log(db: sqlite3.Connection, month: str | None = None, limit: int = 300) -> list[sqlite3.Row]:
    if month:
        return db.execute("SELECT * FROM audit_log WHERE month = ? ORDER BY id DESC LIMIT ?", (month, limit)).fetchall()
    return db.execute("SELECT * FROM audit_log ORDER BY id DESC LIMIT ?", (limit,)).fetchall()


# --------------------------------------------------------------------------
# Запись
# --------------------------------------------------------------------------

def _log(db, user, month, action, employee=None, project=None, old=None, new=None, note=None) -> None:
    db.execute("INSERT INTO audit_log (at, user, month, action, employee, project, old_value, new_value, note) "
               "VALUES (?,?,?,?,?,?,?,?,?)",
               (datetime.now().isoformat(timespec="seconds"), clean(user) or "не указан", month, action, employee, project,
                old, new, note))


def _require_unlocked(db, month) -> None:
    if is_locked(db, month):
        raise MonthLocked(f"месяц {month} закрыт, сначала откройте его заново")


def _employee_name(db, employee_id: int) -> str:
    row = db.execute("SELECT name FROM employees WHERE id = ?", (employee_id,)).fetchone()
    if not row:
        raise ValueError(f"сотрудника {employee_id} нет в базе")
    return row["name"]


def _refresh_total(db, month, employee_id, source: str | None = None) -> None:
    total = db.execute("SELECT COALESCE(SUM(share), 0) FROM allocations WHERE month = ? AND employee_id = ?",
                       (month, employee_id)).fetchone()[0]
    db.execute("UPDATE employee_months SET total_share = ?" + (", source = ?" if source else "") +
               " WHERE month = ? AND employee_id = ?",
               (round(total, 6), *([source] if source else []), month, employee_id))


def save_month(db, month: str, user: str, values: dict[tuple[int, int], float]) -> SaveResult:
    """values: (id сотрудника, id проекта) -> доля от 0 до 1. Ноль означает "нет доли"."""
    check_month(month)
    _require_unlocked(db, month)
    members = {r[0] for r in db.execute("SELECT employee_id FROM employee_months WHERE month = ?", (month,))}
    old: dict[tuple[int, int], float] = {(r["employee_id"], r["project_id"]): r["share"]
                                         for r in db.execute("SELECT employee_id, project_id, share FROM allocations WHERE month = ?", (month,))}
    project_names = {r["id"]: r["name"] for r in db.execute("SELECT id, name FROM projects")}
    touched, changed = set(), 0
    for (eid, pid), raw in values.items():
        if eid not in members or pid not in project_names:
            continue
        if raw < -1e-9 or raw > 1 + 1e-9:
            raise ValidationError([f"доля вне диапазона 0..100%: {raw * 100:g}"])
        new, before = round(max(raw, 0.0), 4), old.get((eid, pid), 0.0)
        if abs(new - before) < 1e-9:
            continue
        if new > 0:
            db.execute("INSERT OR REPLACE INTO allocations (month, employee_id, project_id, share, hours) VALUES (?,?,?,?,NULL)",
                       (month, eid, pid, new))
        else:
            db.execute("DELETE FROM allocations WHERE month = ? AND employee_id = ? AND project_id = ?", (month, eid, pid))
        _log(db, user, month, "изменение", _employee_name(db, eid), project_names[pid], before, new)
        touched.add(eid)
        changed += 1
    for eid in touched:
        _refresh_total(db, month, eid, source="web")
    db.commit()
    return SaveResult(changed, problems(get_month(db, month).rows))


def copy_from_previous(db, month: str, user: str) -> int:
    """Переносит сотрудников и доли прошлого месяца: добавляет отсутствующих и заполняет тех, у кого долей еще нет."""
    check_month(month)
    _require_unlocked(db, month)
    previous = previous_month_with_data(db, month)
    if previous is None:
        raise ValueError("прошлого месяца с данными нет")
    have = {r[0]: r[1] for r in db.execute("SELECT employee_id, total_share FROM employee_months WHERE month = ?", (month,))}
    filled = 0
    for r in db.execute("SELECT employee_id, position, hire_date FROM employee_months WHERE month = ? "
                        "ORDER BY CASE WHEN source_row = 0 THEN 1 ELSE 0 END, source_row, rowid", (previous,)).fetchall():
        eid = r["employee_id"]
        if eid in have and have[eid] > 0:
            continue
        if eid not in have:
            db.execute("INSERT INTO employee_months (month, employee_id, position, hire_date, tag, total_share, is_forecast, source, source_row) "
                       "VALUES (?,?,?,?,NULL,0,0,'web',0)", (month, eid, r["position"], r["hire_date"]))
        db.execute("INSERT OR REPLACE INTO allocations (month, employee_id, project_id, share, hours) "
                   "SELECT ?, employee_id, project_id, share, NULL FROM allocations WHERE month = ? AND employee_id = ?",
                   (month, previous, eid))
        _refresh_total(db, month, eid, source="web")
        filled += 1
    db.execute("INSERT OR IGNORE INTO month_projects (month, project_id, position) "
               "SELECT ?, project_id, position FROM month_projects WHERE month = ?", (month, previous))
    _log(db, user, month, "копирование", note=f"из {previous}, сотрудников: {filled}")
    db.commit()
    return filled


def add_employee(db, month: str, user: str, employee_id: int | None = None, name: str | None = None) -> int:
    check_month(month)
    _require_unlocked(db, month)
    if name and clean(name):
        name = clean(name)
        db.execute("INSERT OR IGNORE INTO employees (name) VALUES (?)", (name,))
        employee_id = db.execute("SELECT id FROM employees WHERE name = ?", (name,)).fetchone()[0]
    if employee_id is None:
        raise ValueError("не указан сотрудник")
    display = _employee_name(db, employee_id)
    if db.execute("SELECT 1 FROM employee_months WHERE month = ? AND employee_id = ?", (month, employee_id)).fetchone():
        raise ValueError(f"{display} уже есть в списке месяца")
    position = db.execute("SELECT position FROM employee_months WHERE employee_id = ? AND position IS NOT NULL AND position != '' "
                          "ORDER BY month DESC LIMIT 1", (employee_id,)).fetchone()
    db.execute("INSERT INTO employee_months (month, employee_id, position, hire_date, tag, total_share, is_forecast, source, source_row) "
               "VALUES (?,?,?,NULL,NULL,0,0,'web',0)", (month, employee_id, position[0] if position else None))
    _log(db, user, month, "добавлен сотрудник", display)
    db.commit()
    return employee_id


def remove_employee(db, month: str, user: str, employee_id: int) -> None:
    check_month(month)
    _require_unlocked(db, month)
    name = _employee_name(db, employee_id)
    total = db.execute("SELECT total_share FROM employee_months WHERE month = ? AND employee_id = ?", (month, employee_id)).fetchone()
    db.execute("DELETE FROM allocations WHERE month = ? AND employee_id = ?", (month, employee_id))
    db.execute("DELETE FROM employee_months WHERE month = ? AND employee_id = ?", (month, employee_id))
    _log(db, user, month, "убран сотрудник", name, old=total[0] if total else None)
    db.commit()


def add_project(db, month: str, user: str, name: str) -> int:
    check_month(month)
    _require_unlocked(db, month)
    name = clean(name or "")
    if not name:
        raise ValueError("не указано название проекта")
    db.execute("INSERT OR IGNORE INTO projects (name) VALUES (?)", (name,))
    pid = db.execute("SELECT id FROM projects WHERE name = ?", (name,)).fetchone()[0]
    db.execute("INSERT OR IGNORE INTO month_projects (month, project_id, position) "
               "VALUES (?, ?, (SELECT COALESCE(MAX(position) + 1, 0) FROM month_projects WHERE month = ?))", (month, pid, month))
    _log(db, user, month, "добавлен проект", project=name)
    db.commit()
    return pid


def lock_month(db, month: str, user: str) -> None:
    """Закрывает месяц, только если у всех сотрудников сумма долей равна 1."""
    check_month(month)
    rows = get_month(db, month).rows
    found = ["в месяце нет сотрудников"] if not rows else problems(rows)
    if found:
        raise ValidationError(found)
    db.execute("INSERT INTO month_status (month, locked, changed_by, changed_at) VALUES (?,1,?,?) "
               "ON CONFLICT(month) DO UPDATE SET locked = 1, changed_by = excluded.changed_by, changed_at = excluded.changed_at",
               (month, clean(user) or "не указан", datetime.now().isoformat(timespec="seconds")))
    _log(db, user, month, "месяц закрыт")
    db.commit()


def unlock_month(db, month: str, user: str) -> None:
    check_month(month)
    db.execute("INSERT INTO month_status (month, locked, changed_by, changed_at) VALUES (?,0,?,?) "
               "ON CONFLICT(month) DO UPDATE SET locked = 0, changed_by = excluded.changed_by, changed_at = excluded.changed_at",
               (month, clean(user) or "не указан", datetime.now().isoformat(timespec="seconds")))
    _log(db, user, month, "месяц открыт заново")
    db.commit()
