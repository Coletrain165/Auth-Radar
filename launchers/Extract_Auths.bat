@echo off
title Pace Auth PDF Extractor
cd /d "%~dp0.."
"venv\Scripts\pythonw.exe" auth_extractor.py
