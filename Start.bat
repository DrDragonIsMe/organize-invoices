@echo off
chcp 65001 >nul
title 发票下载工具
cd /d "%~dp0"
python start.py
if errorlevel 1 (
    echo.
    echo 启动失败，请确保已安装 Python 并加入环境变量
    pause
)
