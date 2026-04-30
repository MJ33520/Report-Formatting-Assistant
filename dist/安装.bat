@echo off
title ReportForge 安装程序

:: 检查管理员权限
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ╔══════════════════════════════════════╗
echo ║     ReportForge v7.0 安装程序       ║
echo ║  政府公文报告格式化工具 - 永久免费  ║
echo ╚══════════════════════════════════════╝
echo.

:: 设置路径
set "INSTALL_DIR=%LOCALAPPDATA%\ReportForge"
set "BIN_DIR=%INSTALL_DIR%\bin"
set "PROFILE_DIR=%INSTALL_DIR%\profiles"
set "SOURCE_DIR=%~dp0"

:: 关闭 Word
echo [1/5] 关闭 Word...
taskkill /f /im WINWORD.EXE >nul 2>&1
timeout /t 2 /nobreak >nul

:: 创建目录
echo [2/5] 创建安装目录...
if not exist "%BIN_DIR%" mkdir "%BIN_DIR%"
if not exist "%PROFILE_DIR%" mkdir "%PROFILE_DIR%"

:: 复制文件
echo [3/5] 复制文件...
copy /y "%SOURCE_DIR%bin\*.dll" "%BIN_DIR%\" >nul
copy /y "%SOURCE_DIR%bin\*.tlb" "%BIN_DIR%\" >nul 2>&1
if exist "%SOURCE_DIR%profiles\default.json" (
    if not exist "%PROFILE_DIR%\default.json" (
        copy /y "%SOURCE_DIR%profiles\default.json" "%PROFILE_DIR%\" >nul
    ) else (
        echo   配置文件已存在，保留用户配置
    )
)
if exist "%SOURCE_DIR%docs" (
    if not exist "%INSTALL_DIR%\docs" mkdir "%INSTALL_DIR%\docs"
    copy /y "%SOURCE_DIR%docs\*.*" "%INSTALL_DIR%\docs\" >nul 2>&1
)

:: RegAsm 注册
echo [4/5] 注册 COM 组件...
set "REGASM=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if not exist "%REGASM%" set "REGASM=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
"%REGASM%" "%BIN_DIR%\ReportForge.AddIn.dll" /codebase /tlb >nul 2>&1
if %errorlevel% neq 0 (
    echo   [!] COM 注册失败，请确保已安装 .NET Framework 4.7.2+
    pause
    exit /b 1
)

:: 写入注册表（Word 自动加载）
echo [5/5] 配置 Word 加载项...
set "REG_KEY=HKCU\Software\Microsoft\Office\Word\Addins\ReportForge.AddIn"
reg add "%REG_KEY%" /v "FriendlyName" /t REG_SZ /d "ReportForge 报告格式化" /f >nul
reg add "%REG_KEY%" /v "Description" /t REG_SZ /d "政府公文报告格式化工具 v7.0" /f >nul
reg add "%REG_KEY%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul

echo.
echo ╔══════════════════════════════════════╗
echo ║         安装完成！                   ║
echo ║                                      ║
echo ║  打开 Word 即可看到「报告格式化」    ║
echo ║  标签页。                            ║
echo ║                                      ║
echo ║  安装目录: %LOCALAPPDATA%\ReportForge
echo ║  配置文件: profiles\default.json     ║
echo ╚══════════════════════════════════════╝
echo.
pause
