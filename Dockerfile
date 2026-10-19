# Аллокация сотрудников по проектам: веб-форма и командные задачи в одном образе.
# Сборка:  docker build -t allocation .
# Если нужен внутренний реестр или зеркало PyPI:
#   docker build --build-arg BASE_IMAGE=registry.example.com/python:3.12-slim \
#                --build-arg PIP_INDEX_URL=https://pypi.example.com/simple -t allocation .
ARG BASE_IMAGE=python:3.12-slim
FROM ${BASE_IMAGE}

ARG PIP_INDEX_URL=""
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONPATH=/app \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir ${PIP_INDEX_URL:+--index-url "$PIP_INDEX_URL"} -r requirements.txt

COPY *.py ./
COPY templates ./templates
COPY examples ./examples
COPY tests ./tests
COPY docker/entrypoint.sh /usr/local/bin/allocation
RUN chmod 0755 /usr/local/bin/allocation \
    && useradd --uid 10001 --create-home --shell /usr/sbin/nologin app \
    && mkdir /data && chown app:app /data

# Рабочая папка с данными. Скрипты читают data/, config/ и пишут out/ относительно нее.
# Монтируйте сюда том: он единственное место, где живут файлы, базы и результаты.
USER app
WORKDIR /data
VOLUME ["/data"]

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import os, urllib.request; urllib.request.urlopen('http://127.0.0.1:%s/healthz' % os.environ.get('WEB_PORT', '8000'), timeout=3)"

ENTRYPOINT ["/usr/local/bin/allocation"]
CMD ["web"]
