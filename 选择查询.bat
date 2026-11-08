@echo off
chcp 65001 >nul
cd /d "%~dp0"

title 选择要查询哪一份成绩

echo ============================================
echo    选择要查询哪一份成绩
echo ============================================
echo.

where py >nul 2>nul
if not errorlevel 1 (
    py -3 queries.py
    goto :done
)
where python >nul 2>nul
if not errorlevel 1 (
    python queries.py
    goto :done
)

echo [错误] 没找到 Python，请先安装 Python 3.8+ 并勾选 Add Python to PATH。
echo.

:done
echo.
pause
