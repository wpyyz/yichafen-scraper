@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

title 易查分成绩批量查询

echo ============================================
echo    易查分 成绩批量查询
echo ============================================
echo.

rem ---------- 1. 找姓名文件 ----------
if not "%~1"=="" (
    set "NAMES=%~1"
    goto :got_names
)
echo 用法：把「姓名.txt」直接拖到这个 bat 文件上就能跑。
echo.
set /p NAMES=或者在这里输入姓名文件路径（可以直接把文件拖进来）: 
if "%NAMES%"=="" (
    echo 没有输入路径，退出。
    pause
    exit /b 1
)
:got_names
set "NAMES=%NAMES:"=%"
if not exist "%NAMES%" (
    echo.
    echo [错误] 找不到文件：%NAMES%
    echo.
    pause
    exit /b 1
)
echo 姓名文件：%NAMES%

rem ---------- 2. 读间隔秒数 ----------
rem 优先级：命令行第 2 个参数 > interval.txt > 12
set "DELAY=%~2"
if not defined DELAY (
    if exist "%~dp0interval.txt" (
        for /f "usebackq delims=" %%L in ("%~dp0interval.txt") do (
            if not defined DELAY set "DELAY=%%L"
        )
    )
)
if not defined DELAY set "DELAY=12"
set "DELAY=%DELAY: =%"
for /f "delims=0123456789" %%a in ("%DELAY%") do set "DELAY=12"
if "%DELAY%"=="" set "DELAY=12"

echo 查询间隔：%DELAY% 秒
echo 想改间隔就双击 设置间隔.bat
echo.

rem ---------- 3. 找 Python ----------
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py -3"
if not defined PY (
    where python >nul 2>nul
    if not errorlevel 1 set "PY=python"
)
if not defined PY (
    echo [错误] 没找到 Python。
    echo.
    echo 请先去 https://www.python.org/downloads/ 装 Python 3.8 或更高版本，
    echo 安装时务必勾选最下面的 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)
echo 使用 %PY%
echo.

rem ---------- 4. 检查依赖 ----------
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo 缺少依赖 openpyxl，正在安装...
    %PY% -m pip install openpyxl
    if errorlevel 1 (
        echo.
        echo [错误] 依赖安装失败。请手动运行：
        echo     %PY% -m pip install openpyxl
        echo.
        pause
        exit /b 1
    )
    echo.
)

rem ---------- 5. 准备输出目录 + 算文件名（不覆盖历史结果） ----------
set "OUTDIR=%~dp0成绩输出"
if not exist "%OUTDIR%" mkdir "%OUTDIR%"
set "IDX=1"
:pick_out
if exist "%OUTDIR%\成绩%IDX%.xlsx" (
    set /a IDX+=1
    goto :pick_out
)
set "OUTFILE=%OUTDIR%\成绩%IDX%.xlsx"
echo 本次结果将保存为：成绩输出\成绩%IDX%.xlsx
echo.

rem ---------- 6. 开跑 ----------
%PY% cli.py --names "%NAMES%" --delay %DELAY% --out "%OUTFILE%" --yes
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo --------------------------------------------------
    echo 没跑成功。上面那段一般会写原因，常见两种：
    echo   1. 找不到 openidCode  ^-^> 先在微信电脑端打开一次那个查询
    echo   2. 想查的不是这一份    ^-^> 加 --uq-code 指定查询码
    echo.
)

echo.
echo ============================================
if exist "%OUTFILE%" (
    echo 完成，结果在：%OUTFILE%
    echo.
    echo 输入 Y 并回车可以打开表格，直接回车则退出。
    set /p OPEN=:
    if /i "!OPEN!"=="Y" start "" "%OUTFILE%"
) else (
    echo 没有生成结果文件，请往上翻看错误信息。
)
echo ============================================
pause
