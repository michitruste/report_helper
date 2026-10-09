@echo off
rem Opens the Report Helper window. Uses, in this order:
rem   - the virtual environment that is active in this terminal
rem   - a .venv / venv / env folder next to this file or one folder up
rem   - "pythonw" / "python" from PATH
cd /d "%~dp0"
set "PYW="
if defined VIRTUAL_ENV if exist "%VIRTUAL_ENV%\Scripts\pythonw.exe" set "PYW=%VIRTUAL_ENV%\Scripts\pythonw.exe"
for %%D in ("%~dp0" "%~dp0..\") do for %%V in (.venv venv env) do (
    if not defined PYW if exist "%%~D%%V\Scripts\pythonw.exe" set "PYW=%%~D%%V\Scripts\pythonw.exe"
)
if defined PYW (
    start "" "%PYW%" "%~dp0report_gui.pyw"
    goto :eof
)
where pythonw >nul 2>nul && (
    start "" pythonw "%~dp0report_gui.pyw"
) || (
    python "%~dp0report_gui.pyw"
    if errorlevel 1 pause
)
