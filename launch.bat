@echo off
setlocal

cd /d "%~dp0"

echo Starting MB Calculator web app...
echo.
echo Open this URL in your browser:
echo http://127.0.0.1:8000/
echo.
echo Press Ctrl+C in this window to stop the server.
echo.

python -m uvicorn web_app:app --host 127.0.0.1 --port 8000

endlocal
