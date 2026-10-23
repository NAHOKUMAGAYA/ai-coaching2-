@echo off
powershell -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dpn0.ps1"
