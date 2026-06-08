@echo off
REM Build Windows GUI executable
cd /d "%~dp0"

echo Installing dependencies...
python -m pip install -r requirements.txt pyinstaller

echo Building OrgChartExtractor.exe ...
python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --name OrgChartExtractor ^
  --add-data "orgchart_defaults.yaml;." ^
  --hidden-import orgchart_crawler ^
  --hidden-import guidelines_index ^
  --hidden-import org_navigator ^
  orgchart_app.py

echo.
echo Copy to the PC aziendale:
echo   dist\OrgChartExtractor.exe
echo   guidelines.xlsx
echo   orgchart_defaults.yaml   (accanto all'exe se serve)
echo.
pause
