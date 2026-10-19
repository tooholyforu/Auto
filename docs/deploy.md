# Развертывание

## 1. Репозиторий в GitLab

Один раз, в папке с кодом:

```sh
ssh-keygen -t ed25519 -C "you@company"             # если ключа еще нет
cat ~/.ssh/id_ed25519.pub                           # вставить в GitLab: Preferences -> SSH Keys
ssh -T git@GITLAB_HOST                              # проверка доступа
git remote add gitlab git@GITLAB_HOST:GROUP/PROJECT.git
git push gitlab claude/jira-confluence-task-mjxean:main    # основная ветка проекта
```

Папки с данными (`data`, `out`, `config`, `work`) и файлы `.env`, `*.xlsx`, `*.sqlite` исключены из репозитория. Перед первой заливкой проверьте: `git ls-files | grep -Ei 'xlsx|sqlite|\.env'` должно быть пустым.

## 2. Сборка образа

```sh
docker build -t allocation .
```

Если сервер сборки не видит интернет, передайте внутренние адреса:

```sh
docker build \
  --build-arg BASE_IMAGE=registry.example.com/library/python:3.12-slim \
  --build-arg PIP_INDEX_URL=https://pypi.example.com/simple \
  -t registry.example.com/team/allocation:1.0 .
docker push registry.example.com/team/allocation:1.0
```

Сертификат корпоративного прокси при сборке: положите его в свою базу образа (`BASE_IMAGE`) или задайте `PIP_CERT`/`PIP_TRUSTED_HOST` в базовом образе. Размер образа около 130 МБ.

Пример `.gitlab-ci.yml` (подстройте под ваши раннеры):

```yaml
stages: [test, build]
test:
  stage: test
  image: python:3.12-slim
  script:
    - pip install -r requirements.txt
    - python -m unittest discover -s tests
build:
  stage: build
  image: docker:27
  services: [docker:27-dind]
  script:
    - docker login -u "$CI_REGISTRY_USER" -p "$CI_REGISTRY_PASSWORD" "$CI_REGISTRY"
    - docker build -t "$CI_REGISTRY_IMAGE:$CI_COMMIT_SHORT_SHA" .
    - docker push "$CI_REGISTRY_IMAGE:$CI_COMMIT_SHORT_SHA"
  rules:
    - if: $CI_COMMIT_BRANCH == $CI_DEFAULT_BRANCH
```

## 3. Запуск на сервере

```sh
mkdir -p work/data/1c work/config work/out
sudo chown -R 10001:10001 work          # либо задайте в .env APP_UID и APP_GID своего пользователя
cp .env.example .env && chmod 600 .env  # заполните значения
docker compose up -d --build
docker compose ps                       # состояние должно быть healthy
curl -s http://127.0.0.1:8000/healthz   # ok
```

В `docker-compose.yml` порт опубликован только на `127.0.0.1`. Наружу открывает обратный прокси (nginx, Traefik) с HTTPS:

```nginx
server {
    listen 443 ssl;
    server_name allocation.example.com;
    # ssl_certificate ...; ssl_certificate_key ...;
    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        # при едином входе на прокси: proxy_set_header X-Forwarded-User $remote_user;
        # и тогда в .env: WEB_USER_HEADER=X-Forwarded-User
        client_max_body_size 20m;
    }
}
```

Если вход выполняет прокси, он обязан затирать заголовок `X-Forwarded-User` из входящих запросов, иначе его можно подделать.

## 4. Периодические задачи

Командные задачи запускаются так же, как в README. Пример cron на сервере (3-го числа в 07:00 сверка за прошлый месяц):

```
0 7 3 * * cd /opt/allocation && docker run --rm --env-file .env --user 10001:10001 -v /opt/allocation/work:/data allocation reconcile --month $(date -d 'last month' +\%Y-\%m)
```

Загрузка аллокации и выгрузок 1С пока зависит от того, кто кладет файлы в `work/data/`. После их появления запускаются `import-allocation`, `import-1c` и `cost-report`.

## 5. Обновление и резервные копии

- Обновление: `git pull && docker compose up -d --build`. Данные в томе `work/` при этом сохраняются.
- Резервное копирование: достаточно копировать папку `work/out/` (файлы `allocation.sqlite` и `costs.sqlite`) и `work/config/`. База SQLite копируется как обычный файл, лучше при остановленном контейнере или командой `sqlite3 файл ".backup копия"`.
- Откат: запустить предыдущий тег образа с тем же томом.

## 6. Что запросить для запуска

| Что | Для чего |
|---|---|
| Служебная учетная запись или токен (PAT) в Jira только на чтение, право Browse Projects на проекты аллокации | `check-jira`, `reconcile`, подсказки в форме |
| Файл сертификата внутреннего удостоверяющего центра | `JIRA_CA_BUNDLE`, проверка TLS без отключения |
| Сетевой доступ с сервера к Jira, к реестру образов GitLab и к зеркалу PyPI (при сборке) | вызовы Jira, сборка и обновление |
| Deploy token GitLab (read_repository, read_registry) для сервера | загрузка кода и образов на сервер |
| Сервер или виртуальная машина с Docker и томом под `work/` (несколько ГБ), резервное копирование | хранение баз и результатов |
| Обратный прокси с HTTPS и, желательно, корпоративным единым входом | безопасный доступ нескольким пользователям |
| Способ получать выгрузки из 1С (папка, HTTP, OData) и ответственный | сейчас файлы кладутся вручную |
