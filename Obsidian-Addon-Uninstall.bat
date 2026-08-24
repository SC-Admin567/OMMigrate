@echo off
title Obsidian Add-on Uninstaller
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] Please right-click and Run this Uninstaller as Administrator!
    pause
    exit /b
)
reg import "%~dp0Obsidian-Addon-UninstallLink.reg"
if exist "%~dp0Obsidian-Addon-Opener.ps1" del /f /q "%~dp0Obsidian-Addon-Opener.ps1"
if exist "%~dp0Obsidian-Addon-SilentRunner.vbs" del /f /q "%~dp0Obsidian-Addon-SilentRunner.vbs"
if exist "%~dp0Obsidian-Addon-DeployLink.reg" del /f /q "%~dp0Obsidian-Addon-DeployLink.reg"
if exist "%~dp0Obsidian-Addon-UninstallLink.reg" del /f /q "%~dp0Obsidian-Addon-UninstallLink.reg"
if exist "%~dp0Obsidian-Addon-Install.bat" del /f /q "%~dp0Obsidian-Addon-Install.bat"
echo Obsidian Add-on has been completely uninstalled.
pause
del "%~f0"
