@echo off
title SLA Report Bot
set "PY=%USERPROFILE%\anaconda3\python.exe"
if not exist "%PY%" set "PY=python"
if "%~1"=="" (
  echo Drag a CSV or Excel ticket file onto SLA_Bot.bat - or paste its path below.
  "%PY%" "%~dp0sla_bot.py"
) else (
  "%PY%" "%~dp0sla_bot.py" "%~1"
)
echo.
pause
