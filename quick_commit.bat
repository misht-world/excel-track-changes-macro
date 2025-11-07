@echo off
cd /d "%~dp0"

git add .
git commit -m "Update: %date% %time%"
git push -f origin main

if %errorlevel% equ 0 (
    echo Quick update completed!
) else (
    echo Update failed!
)
pause