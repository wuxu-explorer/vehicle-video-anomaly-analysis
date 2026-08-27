@echo off
chcp 65001 >nul
title 车辆运营与报警月度报告 V8.4

echo ============================================================
echo        车辆运营与报警月度报告 V8.4
echo        一键生成报告
echo ============================================================
echo.

echo [1/3] 正在检查 Python...
py --version
if errorlevel 1 (
    echo.
    echo [错误] 未检测到 Python。
    echo 请先安装 Python 3.10 或更高版本。
    pause
    exit /b 1
)

echo.
echo [2/3] 正在检查主程序...
if not exist "main.py" (
    echo.
    echo [错误] 未找到 main.py
    echo 当前目录：
    cd
    pause
    exit /b 1
)

echo.
echo [3/3] 开始生成报告...
echo.

py main.py

echo.
echo ============================================================
echo 报告生成程序执行完毕
echo ============================================================
echo.

if errorlevel 1 (
    echo [提示] 程序运行过程中出现错误，请查看上面的错误信息。
) else (
    echo [成功] 报告已经生成。
    echo.
    echo 输出目录：
    echo %CD%\output
)

echo.
pause