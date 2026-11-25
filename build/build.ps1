# Builds dist\TrackChanges.xlam from src\TrackChanges.bas and src\customUI14.xml.
#
# Requirements: Windows, Excel 2010+, and
#   File > Options > Trust Center > Trust Center Settings > Macro Settings >
#   "Trust access to the VBA project object model" turned on (only while building).
#
# Usage:  powershell -ExecutionPolicy Bypass -File build\build.ps1

param(
    [string]$OutFile
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $OutFile) { $OutFile = Join-Path $root 'dist\TrackChanges.xlam' }
$OutFile = [IO.Path]::GetFullPath($OutFile)
$srcBas = Join-Path $root 'src\TrackChanges.bas'
$srcUi  = Join-Path $root 'src\customUI14.xml'

New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) | Out-Null
if (Test-Path $OutFile) { Remove-Item $OutFile -Force }

# The VBA editor imports .bas files in the system ANSI code page.
$tmpDir = Join-Path $env:TEMP ('tc_build_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir | Out-Null
$tmpBas = Join-Path $tmpDir 'TrackChanges.bas'
$code = [IO.File]::ReadAllText($srcBas, [Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
[IO.File]::WriteAllText($tmpBas, $code, [Text.Encoding]::Default)

function Set-DocProp($wb, $name, $value) {
    $props = $wb.BuiltinDocumentProperties
    $prop = [System.__ComObject].InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $props, @($name))
    [System.__ComObject].InvokeMember('Value', [Reflection.BindingFlags]::SetProperty, $null, $prop, @($value)) | Out-Null
}

Write-Host 'Building add-in with Excel...'
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$xl = New-Object -ComObject Excel.Application
$xlPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Add()
    $vbp = $null
    try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count } catch { $vbp = $null }
    if (-not $vbp) {
        throw 'Excel denied access to the VBA project. Turn on "Trust access to the VBA project object model" (File > Options > Trust Center > Trust Center Settings > Macro Settings), run the build, then turn it off again.'
    }
    $null = $vbp.VBComponents.Import($tmpBas)

    # Compile to catch syntax errors (Debug > Compile VBAProject).
    $compile = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($compile -and $compile.Enabled) { $compile.Execute() }

    Set-DocProp $wb 'Title' 'Track Changes'
    Set-DocProp $wb 'Comments' 'Highlights changed cells with conditional formatting. Undo keeps working. https://github.com/misht-world/excel-track-changes-macro'

    $wb.SaveAs($OutFile, 55)   # xlOpenXMLAddIn
    $wb.Close($false)
    $wb = $null
} finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # make sure our own hidden Excel instance is gone (never touches other Excel windows)
    foreach ($id in $xlPid) { Wait-Process -Id $id -Timeout 10 -ErrorAction SilentlyContinue; Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host 'Adding the ribbon tab...'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($OutFile, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $zip.CreateEntry('customUI/customUI14.xml')
    $w = $entry.Open()
    $bytes = [IO.File]::ReadAllBytes($srcUi)
    $w.Write($bytes, 0, $bytes.Length)
    $w.Close()

    $rels = $zip.GetEntry('_rels/.rels')
    $r = New-Object IO.StreamReader($rels.Open())
    $xml = $r.ReadToEnd()
    $r.Close()
    if ($xml -notmatch 'customUI14.xml') {
        $rel = '<Relationship Id="rTcCustomUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
        $xml = $xml -replace '</Relationships>', ($rel + '</Relationships>')
    }
    $rels.Delete()
    $new = $zip.CreateEntry('_rels/.rels')
    $sw = New-Object IO.StreamWriter($new.Open(), (New-Object Text.UTF8Encoding($false)))
    $sw.Write($xml)
    $sw.Close()
} finally {
    $zip.Dispose()
}

Write-Host "Done: $OutFile"
