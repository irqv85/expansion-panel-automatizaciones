<#
Instala Office LTSC 2021 (volumen) solo con Word, Excel y PowerPoint, en espanol.
Excluye todo lo demas: Outlook, Access, Publisher, OneNote, Teams, Lync/Skype,
Groove y OneDrive del instalador de Office.

Uso (PowerShell como administrador):
    .\Instalar-Office-LTSC.ps1
    .\Instalar-Office-LTSC.ps1 -Idioma es-mx
    .\Instalar-Office-LTSC.ps1 -Arquitectura 32
    .\Instalar-Office-LTSC.ps1 -SoloDescargar   (baja los archivos a .\Office)

Usa la Office Deployment Tool (setup.exe). Si no esta junto al script, la baja
del CDN de Microsoft.
#>
param(
    [string]$Idioma = "es-es",
    [ValidateSet("64", "32")][string]$Arquitectura = "64",
    [switch]$SoloDescargar
)

$ErrorActionPreference = "Stop"

# ODT exige elevacion: sin ella setup.exe falla con un error poco claro.
$esAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $esAdmin) {
    Write-Host "Ejecuta este script como administrador." -ForegroundColor Red
    exit 1
}

$trabajo = Join-Path $PSScriptRoot "office-ltsc-setup"
New-Item -ItemType Directory -Force -Path $trabajo | Out-Null

$setup = Join-Path $trabajo "setup.exe"
$config = Join-Path $trabajo "configuration.xml"

# Si ya hay un setup.exe junto al script se reutiliza (instalacion offline).
$local = Join-Path $PSScriptRoot "setup.exe"
if (Test-Path $local) {
    Copy-Item $local $setup -Force
}
elseif (-not (Test-Path $setup)) {
    Write-Host "Descargando Office Deployment Tool..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
        -OutFile $setup -UseBasicParsing
}

# Si ya se bajaron los archivos a .\Office, la instalacion es offline.
$origen = Join-Path $PSScriptRoot "Office"
$sourcePath = ""
if ((Test-Path (Join-Path $origen "Office\Data")) -or $SoloDescargar) {
    $sourcePath = " SourcePath=`"$origen`""
}

# PerpetualVL2021 es el canal de Office LTSC 2021. ProPlus2021Volume trae las
# 3 apps que se piden mas el resto, que se excluye una por una abajo.
$xml = @"
<Configuration>
  <Add OfficeClientEdition="$Arquitectura" Channel="PerpetualVL2021"$sourcePath>
    <Product ID="ProPlus2021Volume">
      <Language ID="$Idioma" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Bing" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="Outlook" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
"@
Set-Content -Path $config -Value $xml -Encoding UTF8

if ($SoloDescargar) {
    Write-Host "Descargando archivos de Office a $origen ..."
    $p = Start-Process -FilePath $setup -ArgumentList "/download `"$config`"" -Wait -PassThru
    if ($p.ExitCode -ne 0) { exit $p.ExitCode }
    Write-Host "Descarga lista." -ForegroundColor Green
    exit 0
}

Write-Host "Instalando Office LTSC 2021 (Word, Excel, PowerPoint, $Idioma)..."
$p = Start-Process -FilePath $setup -ArgumentList "/configure `"$config`"" `
    -Wait -PassThru
if ($p.ExitCode -ne 0) {
    Write-Host "setup.exe termino con codigo $($p.ExitCode)." -ForegroundColor Red
    exit $p.ExitCode
}

Write-Host "Listo. Office LTSC instalado." -ForegroundColor Green
Write-Host "Activacion pendiente: ingresa tu clave de volumen (MAK/KMS) desde Word, Archivo, Cuenta."
