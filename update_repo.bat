@echo off
echo ========================================
echo  Excel Macro Repository Updater
echo ========================================

:: Переходим в папку скрипта
cd /d "%~dp0"

:: Проверяем, есть ли папка .git
if not exist ".git" (
    echo ERROR: This is not a Git repository!
    echo Please run this script from your repository folder.
    pause
    exit /b 1
)

:: Показываем текущие изменения
echo.
echo Current changes:
git status

:: Добавляем все файлы
echo.
echo Adding all files...
git add .

:: Создаем коммит
echo.
set /p commit_message="Enter commit message: "
if "%commit_message%"=="" (
    set commit_message="Auto-update: %date% %time%"
)

git commit -m "%commit_message%"

:: Пушим с принудительной заменой
echo.
echo Uploading to GitHub...
git push -f origin main

:: Результат
echo.
if %errorlevel% equ 0 (
    echo SUCCESS: Repository updated successfully!
) else (
    echo ERROR: Failed to update repository.
)

echo.
pause