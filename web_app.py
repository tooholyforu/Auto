#!/usr/bin/env python3
"""Веб-интерфейс ввода аллокации по месяцам.

Запуск у себя на компьютере:
    python web_app.py --db out/allocation.sqlite
и открыть http://127.0.0.1:8000

Что умеет: показывает сотрудников месяца с долями по проектам, считает сумму на лету, показывает историю
по человеку и подсказки из Jira (где у человека есть задачи), копирует прошлый месяц, закрывает месяц и
ведет журнал изменений. Сервер слушает только локальный адрес и пока без входа по логину: перед публикацией
для других пользователей нужна авторизация.
"""
from __future__ import annotations

import argparse
import math
import os
import secrets
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

from flask import Flask, abort, flash, g, redirect, render_template, request, url_for

import webstore as ws
from allocation_calc import load_roster, month_bounds
from allocation_import import key, load_project_map
from jira_client import load_assignments_csv, load_worklogs_csv
from reconcile_jira import activity_by_project


@dataclass
class Settings:
    db: Path
    roster: Path | None = None            # сотрудники из 1С (для связи с Jira)
    project_map: Path | None = None       # ключи проектов Jira -> проекты аллокации
    jira_dir: Path | None = None          # assignments-ГГГГ-ММ.csv и worklogs-ГГГГ-ММ.csv


def fmt_pct(share) -> str:
    """0.125 -> '12,5', 0.3 -> '30', пусто для нуля."""
    if not share:
        return ""
    text = f"{round(share * 100, 2):.2f}".rstrip("0").rstrip(".")
    return text.replace(".", ",")


def next_month_code(month: str) -> str:
    year, number = map(int, month.split("-"))
    return f"{year + number // 12}-{number % 12 + 1:02d}"


def parse_percent(raw: str) -> float:
    text = (raw or "").strip().replace(",", ".").replace("%", "")
    if not text:
        return 0.0
    value = float(text)
    if not math.isfinite(value) or value < 0 or value > 100:
        raise ValueError(raw)
    return value


def create_app(settings: Settings, secret: str | None = None) -> Flask:
    app = Flask(__name__)
    app.secret_key = secret or os.environ.get("WEB_SECRET") or secrets.token_hex(16)
    app.jinja_env.filters["pct"] = fmt_pct

    def db():
        if "db" not in g:
            g.db = ws.open_db(settings.db)
        return g.db

    @app.teardown_appcontext
    def close_db(_exc):
        conn = g.pop("db", None)
        if conn is not None:
            conn.close()

    @app.before_request
    def same_origin_only():
        if request.method == "POST":
            origin = request.headers.get("Origin")
            if origin and urlparse(origin).netloc != request.host:
                abort(403)

    def user() -> str:
        return (request.form.get("user") or "").strip()

    def back(month: str):
        return redirect(url_for("month_page", month=month))

    def run(month: str, action, ok_message=None):
        """Выполняет действие над месяцем и превращает ошибки в понятные сообщения."""
        try:
            result = action()
        except ws.MonthLocked as exc:
            flash(str(exc), "error")
        except ws.ValidationError as exc:
            flash("Нельзя: " + "; ".join(exc.problems[:8]) + (f" и еще {len(exc.problems) - 8}" if len(exc.problems) > 8 else ""), "error")
        except ValueError as exc:
            flash(str(exc), "error")
        else:
            if ok_message:
                flash(ok_message(result) if callable(ok_message) else ok_message, "ok")
            return result
        return None

    def jira_context(month: str, view: ws.MonthView) -> dict:
        if not (settings.roster and settings.project_map and settings.jira_dir):
            return {"available": False, "note": "Подсказки Jira не подключены (нужны список сотрудников, справочник проектов и папка с данными Jira)."}
        a_path, w_path = settings.jira_dir / f"assignments-{month}.csv", settings.jira_dir / f"worklogs-{month}.csv"
        if not (a_path.exists() or w_path.exists()):
            return {"available": False, "note": f"Данных Jira за {month} нет. Сохраните их командой reconcile_jira.py --save-assignments {a_path}"}
        start, end = month_bounds(month)
        roster = load_roster(settings.roster)
        mapping = load_project_map(settings.project_map)
        assignments = load_assignments_csv(a_path) if a_path.exists() else []
        worklogs = load_worklogs_csv(w_path, start, end) if w_path.exists() else []
        activity, linked = activity_by_project(roster, worklogs, assignments, mapping)
        by_id, by_name = {e.employee_id: e for e in roster}, {key(e.name): e for e in roster}
        cells: dict[str, dict] = {}
        for row in view.rows:
            emp = by_id.get(row.external_id or "") or by_name.get(key(row.name))
            if emp is None or not emp.jira_login:
                continue
            mine = {key(project): data for project, data in activity.get(emp.employee_id, {}).items()}
            for pid, project in view.projects:
                if key(project) in linked:
                    cells[f"{row.employee_id}-{pid}"] = mine.get(key(project), {"tasks": 0, "hours": 0.0})
        return {"available": True, "cells": cells, "tasks": len(assignments), "worklogs": len(worklogs),
                "note": f"Данные Jira за {month}: задач {len(assignments)}, списаний {len(worklogs)}."}

    @app.get("/")
    def index():
        months = ws.list_months(db())
        return back(months[-1] if months else date.today().strftime("%Y-%m"))

    @app.get("/go")
    def go():
        month = request.args.get("month", "")
        try:
            ws.check_month(month)
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))
        return back(month)

    @app.get("/m/<month>")
    def month_page(month):
        try:
            view = ws.get_month(db(), month)
        except ValueError:
            abort(404)
        earlier = [m for m in view.months if m < month]
        later = [m for m in view.months if m > month]
        jira = jira_context(month, view)
        return render_template("month.html", view=view, jira=jira, prev_month=earlier[-1] if earlier else None,
                               next_month=later[0] if later else None, copy_from=view.previous,
                               new_month=next_month_code(max(view.months + [month])))

    @app.post("/m/<month>/save")
    def save(month):
        view = ws.get_month(db(), month)
        values, bad = {}, []
        for row in view.rows:
            for pid, name in view.projects:
                raw = request.form.get(f"v-{row.employee_id}-{pid}")
                if raw is None:
                    continue
                try:
                    values[(row.employee_id, pid)] = parse_percent(raw) / 100
                except ValueError:
                    bad.append(f"{row.name} / {name}: «{raw}»")
        if bad:
            flash("Не сохранено, неверные значения (нужно число от 0 до 100): " + "; ".join(bad[:5]), "error")
            return back(month)
        result = run(month, lambda: ws.save_month(db(), month, user(), values))
        if result is not None:
            flash(f"Сохранено, изменено значений: {result.changed}." if result.changed else "Изменений нет.", "ok")
            if result.problems:
                shown = "; ".join(result.problems[:6]) + (f" и еще {len(result.problems) - 6}" if len(result.problems) > 6 else "")
                flash(f"Сумма долей не равна 100% у {len(result.problems)}: {shown}", "warn")
        return back(month)

    @app.post("/m/<month>/copy")
    def copy(month):
        run(month, lambda: ws.copy_from_previous(db(), month, user()), lambda n: f"Перенесено из прошлого месяца сотрудников: {n}.")
        return back(month)

    @app.post("/m/<month>/add-employee")
    def add_employee(month):
        def action():
            raw_id = request.form.get("employee_id", "")
            return ws.add_employee(db(), month, user(), employee_id=int(raw_id) if raw_id.isdigit() else None,
                                   name=request.form.get("name"))
        run(month, action, "Сотрудник добавлен.")
        return back(month)

    @app.post("/m/<month>/remove-employee")
    def remove_employee(month):
        raw_id = request.form.get("employee_id", "")
        if not raw_id.isdigit():
            abort(400)
        run(month, lambda: ws.remove_employee(db(), month, user(), int(raw_id)), "Сотрудник убран из месяца.")
        return back(month)

    @app.post("/m/<month>/add-project")
    def add_project(month):
        run(month, lambda: ws.add_project(db(), month, user(), request.form.get("name", "")), "Проект добавлен.")
        return back(month)

    @app.post("/m/<month>/lock")
    def lock(month):
        run(month, lambda: ws.lock_month(db(), month, user()), f"Месяц {month} закрыт.")
        return back(month)

    @app.post("/m/<month>/unlock")
    def unlock(month):
        run(month, lambda: ws.unlock_month(db(), month, user()), f"Месяц {month} открыт для правок.")
        return back(month)

    @app.get("/log")
    def log():
        month = request.args.get("month") or None
        return render_template("log.html", entries=ws.recent_log(db(), month), month=month, months=ws.list_months(db()))

    return app


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--db", type=Path, default=Path("out/allocation.sqlite"), help="база аллокации")
    p.add_argument("--roster", type=Path, default=Path("config/employees.csv"), help="сотрудники из 1С (для подсказок Jira)")
    p.add_argument("--project-map", type=Path, default=Path("config/jira_project_map.csv"), help="проекты Jira и аллокации")
    p.add_argument("--jira-dir", type=Path, default=Path("out/jira"), help="папка с assignments-ГГГГ-ММ.csv")
    p.add_argument("--host", default="127.0.0.1", help="адрес (по умолчанию только этот компьютер)")
    p.add_argument("--port", type=int, default=8000)
    args = p.parse_args(argv)
    settings = Settings(args.db, args.roster if args.roster.exists() else None,
                        args.project_map if args.project_map.exists() else None, args.jira_dir)
    app = create_app(settings)
    print(f"База: {args.db}\nОткройте в браузере: http://{args.host}:{args.port}  (остановить: Ctrl+C)")
    app.run(host=args.host, port=args.port, debug=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
