# build.ps1
# Packages the SMA/EMA Basic modules into an installable LibreOffice
# extension (sma_ema.oxt).
#
# It converts each .bas module into LibreOffice's .xba format, generates the
# extension metadata, and zips everything with forward-slash paths (required
# by the .oxt / zip format).
#
# Usage:  pwsh ./build.ps1

$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$libName = 'SmaEma'
$oxtPath = Join-Path $root 'sma_ema.oxt'

# Module name -> source .bas file.
$modules = [ordered]@{
    'MovingAverages'       = Join-Path $root 'MovingAverages.bas'
    'MovingAverageDialog'  = Join-Path $root 'MovingAverageDialog.bas'
}

function ConvertTo-XmlText([string]$s) {
    # Escape the three characters that matter inside XML text content.
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    return $s
}

function New-XbaContent([string]$name, [string]$source) {
    $escaped = ConvertTo-XmlText $source
    return @"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="$name" script:language="StarBasic">$escaped</script:module>
"@
}

# --- Assemble the in-memory list of zip entries (path -> text content) -------
$entries = [ordered]@{}

# Basic modules.
$elementLines = @()
foreach ($name in $modules.Keys) {
    $src = Get-Content -Raw -LiteralPath $modules[$name]
    $entries["$libName/$name.xba"] = New-XbaContent $name $src
    $elementLines += " <library:element library:name=`"$name`"/>"
}

# Module index (script.xlb).
$entries["$libName/script.xlb"] = @"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="$libName" library:readonly="false" library:passwordprotected="false">
$($elementLines -join "`n")
</library:library>
"@

# Dialog index (dialog.xlb). We build the dialog in code rather than storing
# it in the Dialog Editor, so this index is empty -- but LibreOffice still
# requires the file to be present when loading the library.
$entries["$libName/dialog.xlb"] = @"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="$libName" library:readonly="false" library:passwordprotected="false"/>
"@

# Extension description.
$entries['description.xml'] = @"
<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006"
             xmlns:xlink="http://www.w3.org/1999/xlink">
  <identifier value="org.davidjackson.smaema"/>
  <version value="1.0.0"/>
  <display-name>
    <name lang="en">SMA/EMA Moving Averages</name>
  </display-name>
  <publisher>
    <name xlink:href="https://github.com/davidjayjackson/sma_ema" lang="en">David Jackson</name>
  </publisher>
  <extension-description>
    <src lang="en" xlink:href="description/desc_en.txt"/>
  </extension-description>
</description>
"@

$entries['description/desc_en.txt'] =
    'Simple and exponential moving averages for LibreOffice Calc: SMA() and ' +
    'EMA() cell functions plus a dialog for generating result columns.'

# Manifest: register the Basic library folder.
$entries['META-INF/manifest.xml'] = @"
<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">
 <manifest:file-entry manifest:media-type="application/vnd.sun.star.basic-library" manifest:full-path="$libName/"/>
</manifest:manifest>
"@

# --- Write the zip with forward-slash entry names ----------------------------
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (Test-Path $oxtPath) { Remove-Item $oxtPath -Force }

$utf8 = New-Object System.Text.UTF8Encoding($false)
$zip  = [System.IO.Compression.ZipFile]::Open($oxtPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($entryName in $entries.Keys) {
        $entry  = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $stream = $entry.Open()
        $writer = New-Object System.IO.StreamWriter($stream, $utf8)
        $writer.Write([string]$entries[$entryName])
        $writer.Flush()
        $writer.Dispose()
        $stream.Dispose()
    }
}
finally {
    $zip.Dispose()
}

Write-Host "Built $oxtPath"
Write-Host "Entries:"
$entries.Keys | ForEach-Object { Write-Host "  $_" }
