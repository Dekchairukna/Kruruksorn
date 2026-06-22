#!/bin/sh
set -e

# Railway injects PORT at runtime. Keep expansion inside this shell script
# so Gunicorn never receives the literal string "$PORT".
APP_PORT="${PORT:-8080}"
WEB_WORKERS="${WEB_CONCURRENCY:-2}"
WEB_THREADS="${WEB_THREADS:-4}"
WEB_TIMEOUT="${WEB_TIMEOUT:-120}"

echo "Starting KRURUKSORN on port ${APP_PORT}"
exec gunicorn app:app \
  --bind "0.0.0.0:${APP_PORT}" \
  --workers "${WEB_WORKERS}" \
  --threads "${WEB_THREADS}" \
  --timeout "${WEB_TIMEOUT}"
