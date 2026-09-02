@echo off
cd /d "%~dp0"
python schedule_generator.py
if errorlevel 1 (
    echo.
    echo [오류] 근무표 생성에 실패했습니다. 위 메시지를 확인하세요.
    pause
    exit /b 1
)
pause
