# Builds DecodingVBA-Lesson01-VBE-Tour.xlsm for the Beginner VBE tour lesson.
# Requires Microsoft Excel on Windows.
# VBA is included when "Trust access to the VBA project object model" is enabled;
# otherwise the workbook is created with copy-paste code on the Start Here sheet.

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $repoRoot 'public\downloads'
$outFile = Join-Path $outDir 'DecodingVBA-Lesson01-VBE-Tour.xlsm'
$basFile = Join-Path $repoRoot 'workbooks\lesson01\modLesson01.bas'

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$vbaModule = Get-Content -Raw -Path $basFile
# Strip Attribute line — not valid when pasted into CodeModule.AddFromString
$vbaModule = ($vbaModule -replace '(?m)^Attribute VB_Name = ".*"\r?\n', '').Trim()

$excel = $null
$wb = $null
$vbaAdded = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -gt 1) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Start Here'

    $instructions = @(
        'Decoding VBA — Lesson 1: Tour the Visual Basic Editor'
        ''
        'Follow these steps in order:'
        '1. Open the Visual Basic Editor (VBE).'
        '2. In Project Explorer, find this workbook under VBAProject (DecodingVBA-Lesson01-VBE-Tour.xlsm).'
'3. Insert > Module (or open modLesson01 if it already exists).'
'4. If the module is empty, copy the code from the "Code" sheet and paste into the module.'
'5. Click inside HelloDecodingVBA and press F5 to run it.'
'     Click inside ShowActiveSheetName and press F8 to step through (run line by line) the macro.'
'6. Change the MsgBox text in HelloDecodingVBA and run it again.'
'7. Save as .xlsm (macro-enabled). Never save as .xlsx because the macros are removed.'
        ''
        'Next: Option Explicit — https://decodingvba.com/learn/beginner/beginner-option-explicit/'
    )

    for ($i = 0; $i -lt $instructions.Count; $i++) {
        $ws.Cells.Item($i + 1, 1).Value2 = $instructions[$i]
    }

    $ws.Columns.Item('A').ColumnWidth = 100
    $ws.Range('A1').Font.Bold = $true
    $ws.Range('A1').Font.Size = 14

    $codeSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
    $codeSheet.Name = 'Code'
    $codeSheet.Range('A1').Value2 = 'Copy everything below into a new standard module:'
    $codeSheet.Range('A1').Font.Bold = $true
    $codeLines = $vbaModule -split "`r?`n"
    for ($i = 0; $i -lt $codeLines.Count; $i++) {
        $codeSheet.Cells.Item($i + 3, 1).Value2 = $codeLines[$i]
    }
    $codeSheet.Columns.Item('A').ColumnWidth = 110
    $codeSheet.Range('A3:A100').Font.Name = 'Consolas'

    try {
        $vbProj = $wb.VBProject
        if ($null -ne $vbProj) {
            $mod = $vbProj.VBComponents.Add(1)
            $mod.Name = 'modLesson01'
            $mod.CodeModule.AddFromString($vbaModule)
            $vbaAdded = $true
            $ws.Cells.Item($instructions.Count + 2, 1).Value2 =
                'Note: modLesson01 is already in this file — skip the copy step if you see it in Project Explorer.'
        }
    }
    catch {
        Write-Warning "VBA project access denied. Workbook will use copy-paste from the Code sheet."
        Write-Warning "Enable: File > Options > Trust Center > Macro Settings > Trust access to the VBA project object model"
    }

    if (Test-Path $outFile) {
        Remove-Item -Force $outFile
    }

    $wb.SaveAs($outFile, 52)
    Write-Host "Created: $outFile"
    if ($vbaAdded) {
        Write-Host "VBA module modLesson01 embedded."
    }
    else {
        Write-Host 'No embedded VBA - use the Code sheet or import workbooks/lesson01/modLesson01.bas'
    }
}
finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($excel) {
        $excel.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
