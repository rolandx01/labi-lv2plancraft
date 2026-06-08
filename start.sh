#!/bin/bash
# Startet die LV → Plancraft Webapp
set -e
cd "$(dirname "$0")"
source venv/bin/activate
exec python3 app.py
