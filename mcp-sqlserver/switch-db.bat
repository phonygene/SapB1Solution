@echo off
REM MCP SQL Server 資料庫切換腳本 (Windows版本)
REM 用途: 快速切換 MCP Server 連接的資料庫

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ========================================
echo MCP SQL Server 資料庫切換工具
echo ========================================
echo.

REM 檢查當前連接
if exist .env (
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_NAME=" .env') do set CURRENT_DB=%%a
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_SERVER=" .env') do set CURRENT_SERVER=%%a
    echo 目前連接: !CURRENT_SERVER! - !CURRENT_DB!
    echo.
)

echo 可用的資料庫配置:
echo.

set count=0
if exist .env.jtdb (
    set /a count+=1
    echo   [1] jtdb
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_SERVER=" .env.jtdb') do echo       伺服器: %%a
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_NAME=" .env.jtdb') do echo       資料庫: %%a
    echo.
)

if exist .env.JTTST (
    set /a count+=1
    echo   [2] JTTST
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_SERVER=" .env.JTTST') do echo       伺服器: %%a
    for /f "tokens=2 delims==" %%a in ('findstr "^DB_NAME=" .env.JTTST') do echo       資料庫: %%a
    echo.
)

if %count%==0 (
    echo 錯誤: 未找到任何資料庫配置檔案
    pause
    exit /b 1
)

set /p choice="請選擇要切換的資料庫 [1-%count%]: "

if "%choice%"=="1" (
    if exist .env.jtdb (
        copy /y .env.jtdb .env >nul
        echo.
        echo 已成功切換到: jtdb
    ) else (
        echo 錯誤: 配置檔案不存在
        pause
        exit /b 1
    )
) else if "%choice%"=="2" (
    if exist .env.JTTST (
        copy /y .env.JTTST .env >nul
        echo.
        echo 已成功切換到: JTTST
    ) else (
        echo 錯誤: 配置檔案不存在
        pause
        exit /b 1
    )
) else (
    echo 錯誤: 無效的選擇
    pause
    exit /b 1
)

echo.
echo 注意: 請重新啟動 MCP Server 以使更改生效
echo 重啟指令: claude mcp restart
echo.
pause
