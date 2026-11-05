@echo off
chcp 65001 >nul
rem Startet den Converter direkt mit Python, ohne EXE zu bauen.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Erster Start: Umgebung wird eingerichtet ...
    py -3 -m venv .venv 2>nul || python -m venv .venv
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
start "" ".venv\Scripts\pythonw.exe" app.py
