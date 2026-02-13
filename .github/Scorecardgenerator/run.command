#!/bin/bash
set -e
cd "$(dirname "$0")"

# Create venv if it doesn't exist, then install dependencies
if [ ! -d ".venv" ]; then
  python3 -m venv .venv
  source .venv/bin/activate
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt
else
  source .venv/bin/activate
fi

python generate_scorecard.py

# Open the newest PDF if possible
latest_pdf=$(ls -t outputs/*.pdf 2>/dev/null | head -n 1 || true)
if [ -n "$latest_pdf" ]; then
  open "$latest_pdf" || true
fi
