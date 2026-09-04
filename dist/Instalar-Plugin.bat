@echo off
chcp 65001 >nul
title Instalador Metadata Dataverse Document
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-MetadataDataverseDocument.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Ocurrio un error durante la instalacion.
    pause
)
