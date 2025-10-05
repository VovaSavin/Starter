param(
    [string]$PathFileExcel,
    [string]$PathToReport,
    [string]$PathToBackup
)

$officeVersions = "16.0", "15.0", "14.0", "12.0" # Office 365/2019/2016, 2013, 2010, 2007
$isEnabled = $false

foreach ($version in $officeVersions) {
    $path = "HKCU:\Software\Microsoft\Office\$version\Excel\Security"
    if (Test-Path $path) {
        $value = (Get-ItemProperty -Path $path -Name "AccessVBOM" -ErrorAction SilentlyContinue).AccessVBOM
        if ($value -eq 1) {
            Write-Host "✅ Знайдено увімкнений параметр для Office версії $version." -ForegroundColor Green
            $isEnabled = $true
            break # Виходимо з циклу, бо вже знайшли увімкнений параметр
        } else {
            # Встановлюємо значення 1 (увімкнути)
            Set-ItemProperty -Path $path -Name "AccessVBOM" -Value 1 -Type DWord

            Write-Host "Параметр 'Довіряти доступ до об'єктної моделі VBA' було увімкнено."
        }
    }
}

if (-not $isEnabled) {
    Write-Host "❌ Доступ до об'єктної моделі VBA вимкнено для всіх знайдених версій Office." -ForegroundColor Yellow
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true


# path
$filePath = $PathFileExcel

# path to backup
$reportPath = $PathToReport
$backupPath = $PathToBackup

# Transformation path to object of file system
$resolvePath = Resolve-Path -Path $filePath

# Open book
$workbook = $excel.Workbooks.Open($resolvePath.Path)

#  Get access to vba project
$vbProject = $workbook.VBProject

#  Шукаємо рядок з кодом 'Папка_збереження =' для Module3
$module = $vbProject.VBComponents.Item("Module3")
$codeModule = $module.CodeModule

$lines = 1..$codeModule.CountOfLines | ForEach-Object {
    $txt = $codeModule.Lines($_, 1)
    if ($txt -match "Папка_збереження\s*=") { $_ }
}

if ($lines.Count -gt 0) {
    $lastLine = $lines[-1]
    $newCode = "    Папка_збереження = `"$reportPath\`""
    $codeModule.ReplaceLine($lastLine, $newCode)
} else {
    $newCode = "    Папка_збереження = `"$reportPath\`""
    $startLine = 157
    $codeModule.InsertLines($startLine, $newCode)
}


#  Шукаємо рядок з кодом 'folderPath =' для Module1
$module_1 = $vbProject.VBComponents.Item("Module1")
$codeModule_1 = $module_1.CodeModule

$lines = 1..$codeModule_1.CountOfLines | ForEach-Object {
    $txt = $codeModule_1.Lines($_, 1)
    if ($txt -match "folderPath\s*=") { $_ }
}

if ($lines.Count -gt 0) {
    $lastLine = $lines[-1]
    $newCode_1 = "    folderPath = `"$backupPath\`""
    $codeModule_1.ReplaceLine($lastLine, $newCode_1)
} else {
    $newCode_1 = "    folderPath = `"$backupPath\`""
    $startLine_1 = 13
    $codeModule_1.InsertLines($startLine_1, $newCode_1)
}



$workbook.Save()
$excel.Quit()
$excel = $null

# Clean
if ($excel -ne $null) {
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Remove-Variable -Name "excel" -ErrorAction SilentlyContinue
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()