@echo off
chcp 65001 >nul

:: 找最新的 allure-report-* 目录
set "LATEST="
for /f "delims=" %%d in ('dir /b /ad /o-d "allure-report-*" 2^>nul') do (
    if not defined LATEST set "LATEST=%%d"
)

if not defined LATEST (
    echo 未找到 Allure 报告目录，请先执行测试生成报告。
    pause
    exit /b 1
)

echo 打开报告：%LATEST%
"C:\work\tools\allure\allure-2.32.0\bin\allure.bat" open "%LATEST%"
