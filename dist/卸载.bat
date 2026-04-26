@echo off
chcp 65001 >nul
title ReportForge 卸载程序

net session >nul 2>&1
if %errorlevel% neq 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ╔══════════════════════════════════════╗
echo ║     ReportForge 卸载程序            ║
echo ╚══════════════════════════════════════╝
echo.

:: 关闭 Word
echo [1/3] 关闭 Word...
taskkill /f /im WINWORD.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

:: 取消 COM 注册
echo [2/3] 取消 COM 注册...
set "BIN_DIR=%LOCALAPPDATA%\ReportForge\bin"
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM%" set "REGASM=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
if exist "%BIN_DIR%\ReportForge.AddIn.dll" (
    "%REGASM%" "%BIN_DIR%\ReportForge.AddIn.dll" /unregister >nul 2>&1
)

:: 删除注册表
echo [3/3] 清理注册表...
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\ReportForge.AddIn" /f >nul 2>&1

echo.
echo 卸载完成！
echo.
echo 您的配置文件保留在:
echo   %LOCALAPPDATA%\ReportForge\profiles\
echo.
echo 如需完全清除，请手动删除:
echo   %LOCALAPPDATA%\ReportForge\
echo.
pause
