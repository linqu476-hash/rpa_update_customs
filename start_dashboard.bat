@echo off
rem ---------------------------------------------------------------
rem  Export invoice automation - one-click start (run on a PC that is
rem  inside the company network, or on VPN, so it can reach D365)
rem ---------------------------------------------------------------
cd /d "%~dp0"

if not exist ".env" (
    echo.
    echo  The file .env is missing. Create it in this folder with these lines:
    echo      ERP_USER=your_username@dc.libgroup.com
    echo      ERP_PASSWORD=your_password
    echo      APP_PASSWORD=choose-a-dashboard-password
    echo.
    pause
    exit /b 1
)

where py >nul 2>nul
if errorlevel 1 (
    echo  Python was not found. Install it from https://www.python.org/downloads/windows/
    echo  and tick "Add python.exe to PATH" in the installer.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo  First start: setting up (a few minutes, once only)...
    py -m venv .venv
)
call ".venv\Scripts\activate.bat"
python -m pip install --quiet --disable-pip-version-check -r requirements.txt

echo.
echo  Dashboard is starting. Colleagues on the company network can open:
echo      http://%COMPUTERNAME%:8501
echo  Keep this window open while the dashboard is in use.
echo.
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
