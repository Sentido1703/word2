#!/bin/bash
set -e

export ADMIN_USER_MODEL=AdminUser
export ADMIN_USER_MODEL_USERNAME_FIELD=username
export ADMIN_SECRET_KEY=secret_key

echo "Waiting for PostgreSQL..."
while ! nc -z db 5432; do
  sleep 1
done
echo "PostgreSQL started"

echo "Initializing database..."
python init_db.py

echo "Starting application..."
exec uvicorn main:app --host 0.0.0.0 --port 8000
