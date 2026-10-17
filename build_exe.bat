@echo off
REM Builds dist\DUD-Converter.exe (a single file, no Python needed to run it).
python -m pip install -r requirements-dev.txt || goto :error
python -m PyInstaller --noconfirm --onefile --windowed --name DUD-Converter app.py || goto :error
echo.
echo Done: dist\DUD-Converter.exe
goto :eof
:error
echo Build failed.
exit /b 1
