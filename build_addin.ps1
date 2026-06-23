# build_addin.ps1
# Builds the MOVAVG Python Calc Add-In into an installable LibreOffice
# extension (sma_ema_addin.oxt).
#
# Steps:
#   1. Compile addin/idl/MovingAverages.idl into a UNO type library
#      (MovingAverages.rdb) with the SDK's unoidl-write.
#   2. Zip the add-in files (Python component, type library, Function
#      Wizard config, metadata) with forward-slash paths into the .oxt.
#
# Requires LibreOffice + its SDK installed. Override the location with
# -LoRoot if LibreOffice is not at the default path.
#
# Usage:  pwsh ./build_addin.ps1 [-LoRoot "C:\Program Files\LibreOffice"]

param(
    [string]$LoRoot = 'C:\Program Files\LibreOffice'
)

$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$addin   = Join-Path $root 'addin'
$oxtPath = Join-Path $root 'sma_ema_addin.oxt'

$program = Join-Path $LoRoot 'program'
$sdkBin  = Join-Path $LoRoot 'sdk\bin'
$unoidl  = Join-Path $sdkBin 'unoidl-write.exe'
$typesRdb = Join-Path $program 'types.rdb'

foreach ($p in @($unoidl, $typesRdb)) {
    if (-not (Test-Path $p)) {
        throw "Required file not found: $p`nIs LibreOffice + the SDK installed at '$LoRoot'? Pass -LoRoot to override."
    }
}

# --- 1. Compile the IDL -> type library -------------------------------------
# unoidl-write needs LibreOffice's program dir on PATH to find its DLLs.
$env:PATH = "$program;$env:PATH"

$idl    = Join-Path $addin 'idl\MovingAverages.idl'
$rdbOut = Join-Path $addin 'MovingAverages.rdb'
if (Test-Path $rdbOut) { Remove-Item $rdbOut -Force }

Write-Host "Compiling $idl ..."
& $unoidl $typesRdb $idl $rdbOut
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $rdbOut)) {
    throw "unoidl-write failed (exit $LASTEXITCODE)."
}
Write-Host "  -> $rdbOut"

# --- 2. Package the .oxt -----------------------------------------------------
# entry-name-inside-oxt -> source file on disk
$files = [ordered]@{
    'MovingAverages.rdb'      = $rdbOut
    'MovingAveragesAddIn.py'  = Join-Path $addin 'MovingAveragesAddIn.py'
    'CalcAddIns.xcu'          = Join-Path $addin 'CalcAddIns.xcu'
    'description.xml'         = Join-Path $addin 'description.xml'
    'description/desc_en.txt' = Join-Path $addin 'description\desc_en.txt'
    'META-INF/manifest.xml'   = Join-Path $addin 'META-INF\manifest.xml'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (Test-Path $oxtPath) { Remove-Item $oxtPath -Force }

$zip = [System.IO.Compression.ZipFile]::Open($oxtPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($entryName in $files.Keys) {
        $src = $files[$entryName]
        if (-not (Test-Path $src)) { throw "Missing file: $src" }
        # CreateEntryFromFile keeps the exact bytes (works for the binary
        # .rdb) and we pass the forward-slash entry name explicitly.
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $src, $entryName,
            [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
    }
}
finally {
    $zip.Dispose()
}

Write-Host "Built $oxtPath"
$files.Keys | ForEach-Object { Write-Host "  $_" }
