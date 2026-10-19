#!/bin/sh
# Точка входа контейнера: allocation <команда> [параметры]
set -eu

usage() {
cat <<'TXT'
Использование: allocation <команда> [параметры]

Команды:
  web                 веб-форма ввода аллокации (по умолчанию), порт 8000
  import-allocation   загрузка аллокации из Excel в базу       (allocation_import.py)
  import-1c           загрузка выгрузок 1С с расходами         (onec_import.py)
  cost-report         расходы по проектам за месяц, Excel      (cost_report.py)
  reconcile           сверка аллокации с задачами Jira         (reconcile_jira.py)
  check-jira          проверка доступа к Jira                  (check_jira.py)
  sync-jira           расчет долей из часов Jira (необязательно) (sync_jira.py)
  test                тесты
  help                эта справка

Параметры команды передаются дальше, например:
  allocation import-1c --month 2026-09 --dir data/1c
Справка по параметрам команды: allocation <команда> --help
Рабочая папка /data (том): data/ входные файлы, config/ справочники, out/ результаты.
TXT
}

cmd="${1:-web}"
[ "$#" -gt 0 ] && shift

case "$cmd" in
  web)
    exec python -m waitress --listen="${WEB_HOST:-0.0.0.0}:${WEB_PORT:-8000}" \
      --threads="${WEB_THREADS:-4}" --call web_app:create_app_from_env ;;
  import-allocation) exec python /app/allocation_import.py "$@" ;;
  import-1c)         exec python /app/onec_import.py "$@" ;;
  cost-report)       exec python /app/cost_report.py "$@" ;;
  reconcile)         exec python /app/reconcile_jira.py "$@" ;;
  check-jira)        exec python /app/check_jira.py "$@" ;;
  sync-jira)         exec python /app/sync_jira.py "$@" ;;
  test)              exec python -m unittest discover -s /app/tests ;;
  help|-h|--help)    usage ;;
  *)                 echo "Неизвестная команда: $cmd" >&2; echo >&2; usage >&2; exit 2 ;;
esac
