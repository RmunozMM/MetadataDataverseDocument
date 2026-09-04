# Install-MetadataDataverseDocument.ps1
# Script de instalacion automatica para XrmToolBox
$ErrorActionPreference = 'Stop'

Clear-Host
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Instalador de Metadata Dataverse Document para XrmToolBox " -ForegroundColor Cyan
Write-Host "  Version: 2.1.0.0 (Release)                              " -ForegroundColor Cyan
Write-Host "  Desarrollador: Rogelio Munoz (www.rogeliomunoz.cl)       " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

$pluginsDir = Join-Path $env:APPDATA "MscrmTools\XrmToolBox\Plugins"

if (-not (Test-Path $pluginsDir)) {
    Write-Host "[1/4] Creando directorio de plugins de XrmToolBox..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $pluginsDir -Force | Out-Null
    Write-Host "      Directorio creado: $pluginsDir" -ForegroundColor Gray
} else {
    Write-Host "[1/4] Directorio de XrmToolBox verificado." -ForegroundColor Green
}

$subfolder = Join-Path $pluginsDir "MetadataDataverseDocument"
if (-not (Test-Path $subfolder)) {
    New-Item -ItemType Directory -Path $subfolder -Force | Out-Null
}

$currentDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($currentDir)) {
    $currentDir = (Get-Location).Path
}

# Desbloquear archivos
Write-Host "[2/4] Desbloqueando archivos DLL descargados..." -ForegroundColor Green
$unblockedCount = 0
Get-ChildItem -Path $currentDir -Filter *.dll -Recurse | ForEach-Object {
    Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
    $unblockedCount++
}
Write-Host "      $unblockedCount archivo(s) DLL desbloqueado(s)." -ForegroundColor Gray

# Copiar DLL principal
Write-Host "[3/4] Instalando plugin principal..." -ForegroundColor Green
$mainDll = Join-Path $currentDir "MetadataDataverseDocument.dll"
if (Test-Path $mainDll) {
    Copy-Item -Path $mainDll -Destination $pluginsDir -Force
    Write-Host "      [OK] MetadataDataverseDocument.dll instalado en Plugins" -ForegroundColor White
} else {
    Write-Host "ERROR: No se encontro MetadataDataverseDocument.dll en $currentDir" -ForegroundColor Red
    Write-Host ""
    Write-Host "Presione Enter para salir..." -ForegroundColor Yellow
    Read-Host
    Exit 1
}

# Copiar dependencias a subcarpeta y directorio raiz
Write-Host "[4/4] Instalando librerias y dependencias..." -ForegroundColor Green
$deps = @("EPPlus.dll", "System.Resources.Extensions.dll")
foreach ($dep in $deps) {
    $depPath = Join-Path $currentDir $dep
    if (Test-Path $depPath) {
        Copy-Item -Path $depPath -Destination $subfolder -Force
        Copy-Item -Path $depPath -Destination $pluginsDir -Force
        Write-Host "      [OK] $dep instalado" -ForegroundColor White
    }
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "         INSTALACION COMPLETADA EXITOSAMENTE!             " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  1. Abra su aplicacion XrmToolBox." -ForegroundColor White
Write-Host "  2. En la lista de plugins busque: 'Metadata Dataverse Document'." -ForegroundColor White
Write-Host "  3. Desarrollado por Rogelio Munoz (www.rogeliomunoz.cl)." -ForegroundColor White
Write-Host ""
Write-Host "Presione Enter para cerrar esta ventana..." -ForegroundColor Cyan
Read-Host