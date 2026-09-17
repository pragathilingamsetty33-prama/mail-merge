@echo off
title Certificate Generation Agent
color 0F
cd /d "%~dp0"
echo ================================================================
echo         Launching Certificate Generation Agent...
echo ================================================================
python certificate_agent.py
echo.
pause
