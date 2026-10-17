@echo off
setlocal
rem ============================================================
rem  校园网自动登录 —— 双击入口
rem
rem  优先在 Windows Terminal 里运行(菜单的背景图功能需要 WT)。
rem  这台机器的 Windows Terminal COM 委托注册缺失,系统会静默回退到
rem  conhost,所以这里直接调 wt.exe,绕开委托。
rem
rem  用法:
rem    connect.bat                打开菜单
rem    connect.bat -Mode status   只看在线状态
rem    connect.bat -Mode doctor   一键诊断
rem ============================================================

set "SELF=%~dp0Connect-CampusNet.ps1"

where wt.exe >nul 2>nul
if not errorlevel 1 (
    rem -d 指定起始目录;末尾的 . 是避开 "路径\" 里反斜杠转义引号的老坑
    start "" wt.exe -d "%~dp0." powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SELF%" %*
    exit /b
)

rem 没装 Windows Terminal:退回普通控制台(背景图不可用)
echo [提示] 没找到 Windows Terminal,用普通控制台启动。
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SELF%" %*
if "%~1"=="" pause
