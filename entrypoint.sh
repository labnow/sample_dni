#!/bin/bash
mv /app/persist_old /app/persist
python3 -m flask run --host=0.0.0.0

exec "$@"