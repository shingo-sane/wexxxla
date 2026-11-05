@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Wexxxla - EXE erstellen
cd /d "%~dp0"

echo.
echo  ==========================================================
echo     Wexxxla  -  EXE wird erstellt
echo  ==========================================================
echo.

rem ---------- 1. Python suchen (oder installieren) ----------
set "PY="
where py >nul 2>nul && (py -3 -c "import sys" >nul 2>nul && set "PY=py -3")
if not defined PY (
    where python >nul 2>nul && (python -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>nul && set "PY=python")
)
if not defined PY (
    for %%V in (313 312 311 310) do (
        if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
    )
)
if not defined PY (
    echo  Python ist nicht installiert.
    where winget >nul 2>nul
    if errorlevel 1 (
        echo  Bitte installiere Python 3.12 von https://www.python.org/downloads/
        echo  WICHTIG: Beim Installieren "Add python.exe to PATH" anhaken.
        echo  Danach diese Datei erneut starten.
        pause
        exit /b 1
    )
    echo  Python 3.12 wird jetzt automatisch installiert ...
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set PY="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    ) else (
        echo  Python wurde installiert. Bitte diese Datei noch einmal starten.
        pause
        exit /b 1
    )
)
echo  [1/4] Python gefunden: !PY!

rem ---------- 2. Eigene Umgebung + Bibliotheken ----------
if not exist ".venv\Scripts\python.exe" (
    echo  [2/4] Erstelle Arbeitsumgebung ...
    !PY! -m venv .venv
    if errorlevel 1 ( echo  Fehler beim Erstellen der Umgebung. & pause & exit /b 1 )
)
set "VPY=.venv\Scripts\python.exe"
echo  [2/4] Lade Bibliotheken herunter (beim ersten Mal einige Minuten) ...
"%VPY%" -m pip install --upgrade pip >nul
"%VPY%" -m pip install -r requirements.txt pyinstaller
if errorlevel 1 ( echo. & echo  Fehler beim Installieren der Bibliotheken. Internetverbindung pruefen. & pause & exit /b 1 )

rem ---------- 3. EXE bauen ----------
echo  [3/4] Baue Wexxxla.exe ...
"%VPY%" build.py
if errorlevel 1 ( echo. & echo  Der Build ist fehlgeschlagen - siehe Meldungen oben. & pause & exit /b 1 )

rem ---------- 4. Verknuepfung auf dem Desktop ----------
echo  [4/4] Lege Desktop-Verknuepfung an ...
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Wexxxla.lnk'); $s.TargetPath='%CD%\dist\Wexxxla\Wexxxla.exe'; $s.WorkingDirectory='%CD%\dist\Wexxxla'; $s.IconLocation='%CD%\dist\Wexxxla\Wexxxla.exe,0'; $s.Save()"

echo.
echo  ==========================================================
echo   Fertig!
echo   Die App liegt hier:  %CD%\dist\Wexxxla\Wexxxla.exe
echo   Auf dem Desktop gibt es jetzt die Verknuepfung "Wexxxla".
echo   Den Ordner dist\Wexxxla kannst du beliebig verschieben.
echo  ==========================================================
echo.

rem ---------- Office-Programm pruefen ----------
reg query "HKCR\Word.Application" >nul 2>nul && goto :haveoffice
if exist "%ProgramFiles%\LibreOffice\program\soffice.exe" goto :haveoffice
if exist "%ProgramFiles(x86)%\LibreOffice\program\soffice.exe" goto :haveoffice
echo  Hinweis: Weder Microsoft Office noch LibreOffice wurde gefunden.
echo  Ohne eines davon fehlen z.B. Word/Excel/PowerPoint zu PDF.
where winget >nul 2>nul || goto :nooffice
choice /C JN /M "  LibreOffice (kostenlos) jetzt installieren"
if errorlevel 2 goto :nooffice
winget install -e --id TheDocumentFoundation.LibreOffice --accept-package-agreements --accept-source-agreements
goto :haveoffice
:nooffice
echo  Download: https://de.libreoffice.org/download/download/
:haveoffice

choice /C JN /M "  App jetzt starten"
if errorlevel 2 exit /b 0
start "" "dist\Wexxxla\Wexxxla.exe"
exit /b 0
