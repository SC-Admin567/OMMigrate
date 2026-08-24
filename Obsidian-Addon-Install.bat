@echo off
title Obsidian Add-on Installer
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Please right-click and Run this Installer as Administrator!
    pause
    exit /b
)
reg import "%~dp0Obsidian-Addon-DeployLink.reg"
echo Installation Complete! The Obsidian Add-on is successfully activated.
pause
