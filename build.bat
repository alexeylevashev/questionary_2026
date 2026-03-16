@echo off
echo ============================================
echo  Сборка questionary.exe
echo ============================================

REM Установить PyInstaller если нет
python -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Устанавливаю PyInstaller...
    python -m pip install pyinstaller
)

REM Очистить предыдущую сборку
if exist "dist\questionary" rmdir /s /q "dist\questionary"
if exist "build\questionary" rmdir /s /q "build\questionary"

REM Собрать
echo.
echo Запускаю PyInstaller...
python -m PyInstaller questionary.spec

if %errorlevel% neq 0 (
    echo.
    echo ОШИБКА сборки! Смотри лог выше.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Готово! Папка с exe: dist\questionary\
echo  Запуск: dist\questionary\questionary.exe
echo ============================================
pause
