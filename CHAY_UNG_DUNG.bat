@echo off
cd /d "%~dp0"
py xet_tot_nghiep.py
if errorlevel 1 pause
