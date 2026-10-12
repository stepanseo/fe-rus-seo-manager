@echo off
chcp 65001 >nul
echo ==========================================
echo Установка WordKeeper TOP-10
echo ==========================================
echo.
echo [1/2] Устанавливаю Python-пакет Playwright...
py -m pip install playwright
if errorlevel 1 (
    echo.
    echo Не удалось через "py". Пробую через "python"...
    python -m pip install playwright
    if errorlevel 1 (
        echo.
        echo ОШИБКА: не удалось установить Playwright.
        pause
        exit /b 1
    )
)

echo.
echo [2/2] Устанавливаю браузер Chromium для Playwright...
py -m playwright install chromium
if errorlevel 1 (
    echo.
    echo Не удалось через "py". Пробую через "python"...
    python -m playwright install chromium
    if errorlevel 1 (
        echo.
        echo ОШИБКА: не удалось установить Chromium.
        pause
        exit /b 1
    )
)

echo.
echo ==========================================
echo ГОТОВО. Playwright + Chromium установлены.
echo Теперь запустите FE-RUS SEO Manager TOP10.
echo ==========================================
pause
