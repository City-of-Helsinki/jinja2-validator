#!/bin/bash
set -e

# If a command is passed, run that
if [[ -n "$*" ]]; then
  exec "$@"
fi

# Default Gunicorn settings (can be overridden via env)
: "${PORT:=5001}"
: "${WORKERS:=4}"
: "${THREADS:=1}"
: "${TIMEOUT:=60}"

echo "Starting Gunicorn on port ${PORT} (workers=${WORKERS}, threads=${THREADS})..."

exec gunicorn \
  --bind "0.0.0.0:${PORT}" \
  --workers "${WORKERS}" \
  --threads "${THREADS}" \
  --timeout "${TIMEOUT}" \
  --access-logfile '-' \
  --error-logfile '-' \
  validator_api:app
