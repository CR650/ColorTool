@echo off
title 颜色主题管理工具 - HTTP服务器
echo ========================================
echo    颜色主题管理工具 - HTTP服务器
echo ========================================
echo.
echo 正在启动服务器...
echo 服务器地址: http://localhost:8000
echo.
echo 提示：
echo - 服务器启动后会自动打开浏览器
echo - 按 Ctrl+C 可停止服务器
echo - 关闭此窗口也会停止服务器
echo.
echo ========================================

:: 启动服务器并自动打开浏览器
start http://localhost:8000
python -m http.server 8000

echo.
echo 服务器已停止
pause
