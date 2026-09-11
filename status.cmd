@echo off
title Safripol Operations KPI Report - Spot Check
set "REPO=%~dp0"
powershell -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%REPO%tools\status.ps1" %*
