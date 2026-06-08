@echo off
cd /d "%~dp0"

if not exist config.yaml (
    echo config.yaml not found.
    echo Copy config.example.yaml to config.yaml and fill in your tokens.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt -q
python orgchart_crawler.py
pause
