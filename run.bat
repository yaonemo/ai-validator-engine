@echo off
chcp 65001 >nul
echo ========================================
echo   通用校验引擎 - 一键运行
echo ========================================
echo.
echo 用法：把要校验的 Excel 拖到这个窗口，按回车
echo 不拖文件直接回车 = 用示例数据演示
echo.
set /p FILE=请输入数据文件路径（或直接回车）: 
if "%FILE%"=="" (
    python run.py
) else (
    python run.py "%FILE%"
)
echo.
pause
