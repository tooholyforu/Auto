"""Клиент Jira Server / Data Center (REST API v2): списания времени за период.

Доступ: персональный токен (Jira 8.14+, заголовок Bearer) или логин и пароль.
Нужны права просмотра проектов и списаний времени. Данные только читаются.
"""
from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterator

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

DELIM = ";"
WORKLOG_FIELDS = ["worklog_id", "issue_key", "project_key", "login", "user_key", "email", "started", "seconds"]


class JiraError(RuntimeError):
    pass


@dataclass(frozen=True)
class Worklog:
    worklog_id: str
    issue_key: str
    project_key: str
    login: str          # author.name, логин в Jira Server
    user_key: str       # author.key
    email: str
    started: date
    seconds: int


def explain(exc: Exception) -> str:
    """Понятное сообщение вместо длинного текста исключения при сбоях сети."""
    if isinstance(exc, JiraError):
        return str(exc)
    if isinstance(exc, requests.exceptions.SSLError):
        return "ошибка проверки сертификата: укажите файл сертификата вашего центра в JIRA_CA_BUNDLE"
    if isinstance(exc, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
        return "сервер Jira недоступен: проверьте, что включен VPN, и адрес в JIRA_URL"
    return f"запрос к Jira не удался: {exc}"


def build_session(token: str | None = None, user: str | None = None, password: str | None = None,
                  ca_bundle: str | None = None) -> requests.Session:
    """Сессия с повторами при 429/5xx. Проверка сертификата не отключается: для внутреннего
    удостоверяющего центра передайте путь к его сертификату в ca_bundle."""
    session = requests.Session()
    if token:
        session.headers["Authorization"] = f"Bearer {token}"
    elif user and password:
        session.auth = (user, password)
    else:
        raise JiraError("не заданы реквизиты доступа: нужен JIRA_TOKEN или JIRA_USER и JIRA_PASSWORD")
    session.headers["Accept"] = "application/json"
    session.verify = ca_bundle or True
    retry = Retry(total=3, backoff_factor=1, status_forcelist=(429, 502, 503, 504),
                  allowed_methods=("GET",), respect_retry_after_header=True)
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    return session


class JiraClient:
    def __init__(self, base_url: str, session, timeout: int = 30, page_size: int = 100):
        self.base = base_url.rstrip("/")
        self.session = session
        self.timeout = timeout
        self.page_size = page_size

    def _get(self, path: str, **params) -> dict:
        response = self.session.get(f"{self.base}{path}", params=params, timeout=self.timeout)
        if response.status_code in (401, 403):
            raise JiraError(f"Jira отказала в доступе ({response.status_code}): проверьте токен и права "
                            "на просмотр проектов и списаний времени")
        response.raise_for_status()
        return response.json()

    def myself(self) -> dict:
        return self._get("/rest/api/2/myself")

    def projects(self) -> list[dict]:
        return self._get("/rest/api/2/project")

    def worklogs(self, start: date, end: date) -> list[Worklog]:
        """Списания с датой начала от start включительно до end не включая."""
        return list(self._iter_worklogs(start, end))

    def _iter_worklogs(self, start: date, end: date) -> Iterator[Worklog]:
        jql = f'worklogDate >= "{start:%Y-%m-%d}" AND worklogDate < "{end:%Y-%m-%d}"'
        offset = 0
        while True:
            page = self._get("/rest/api/2/search", jql=jql, fields="project,worklog",
                             startAt=offset, maxResults=self.page_size)
            issues = page.get("issues", [])
            for issue in issues:
                yield from self._issue_worklogs(issue, start, end)
            offset += len(issues)
            if not issues or offset >= page.get("total", 0):
                return

    def _issue_worklogs(self, issue: dict, start: date, end: date) -> Iterator[Worklog]:
        key = issue["key"]
        fields = issue.get("fields") or {}
        project_key = (fields.get("project") or {}).get("key") or key.rsplit("-", 1)[0]
        block = fields.get("worklog") or {}
        entries = block.get("worklogs", [])
        if block.get("total", len(entries)) > len(entries):
            entries = self._all_worklogs(key)       # в ответе поиска списания обрезаны
        for entry in entries:
            started = date.fromisoformat(entry["started"][:10])
            if not start <= started < end:
                continue
            author = entry.get("author") or {}
            yield Worklog(str(entry.get("id", "")), key, project_key, author.get("name") or "",
                          author.get("key") or "", author.get("emailAddress") or "",
                          started, int(entry.get("timeSpentSeconds", 0)))

    def _all_worklogs(self, issue_key: str) -> list[dict]:
        entries: list[dict] = []
        while True:
            page = self._get(f"/rest/api/2/issue/{issue_key}/worklog", startAt=len(entries), maxResults=1000)
            chunk = page.get("worklogs", [])
            entries += chunk
            if not chunk or len(entries) >= page.get("total", 0):
                return entries


def save_worklogs_csv(worklogs: list[Worklog], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=DELIM)
        writer.writerow(WORKLOG_FIELDS)
        for w in worklogs:
            writer.writerow([w.worklog_id, w.issue_key, w.project_key, w.login, w.user_key, w.email,
                             w.started.isoformat(), w.seconds])


def load_worklogs_csv(path: Path, start: date, end: date) -> list[Worklog]:
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=DELIM))
    missing = set(WORKLOG_FIELDS) - set(rows[0] if rows else WORKLOG_FIELDS)
    if missing:
        raise SystemExit(f"{path}: нет столбцов {', '.join(sorted(missing))}")
    result = [Worklog(r["worklog_id"], r["issue_key"], r["project_key"], r["login"], r["user_key"], r["email"],
                      date.fromisoformat(r["started"]), int(r["seconds"])) for r in rows]
    return [w for w in result if start <= w.started < end]
