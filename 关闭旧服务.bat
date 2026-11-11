@echo off
echo Closing old service on port 8766 ...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8766 ^| findstr LISTENING') do (
    echo Killing process PID %%a
    taskkill /F /PID %%a
)
echo Done. Now double-click start_ui.bat to restart.
pause