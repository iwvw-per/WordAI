@echo off
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" /f >nul 2>&1
echo ===================================================
echo [SUCCESS] Office Event-Based Debug Registry Cleaned.
echo ===================================================
pause
