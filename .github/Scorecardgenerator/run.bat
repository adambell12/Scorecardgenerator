@echo off
setlocal
cd /d %~dp0

if not exist ".venv\" (
  python -m venv .venv
  call .venv\Scripts\activate.bat
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt
) else (
  call .venv\Scripts\activate.bat
)

python generate_scorecard.py

for %%F in ("outputs\Scorecard - *.pdf") do set "LASTPDF=%%F"
if defined LASTPDF start "" "%LASTPDF%"
endlocal
