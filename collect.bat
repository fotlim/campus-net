@echo off
chcp 936 >nul
cd /d "%~dp0"
echo ================================================
echo   校园网诊断采集(约 20 秒)
echo   采集完文字会自动复制到剪贴板,直接粘贴发出去
echo ================================================
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Connect-CampusNet.ps1" -Mode collect
echo.
echo 如果同目录下出现了 collect-日期.zip,把它也一起发出去。
pause