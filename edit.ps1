$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true


# path
$filePath = "V:\test\Обробка витрат ВП_08_1.2025\Обробка витрат ВП_08_1.20252.xlsm"

# Transformation path to object of file system
$resolvePath = Resolve-Path -Path $filePath

# Open book
$workbook = $excel.Workbooks.Open($resolvePath.Path)



#  Get access to vba project
$vbProject = $workbook.VBProject
$module = $vbProject.VBComponents.Item("Module3")
$codeModule = $module.CodeModule

# Add needed code
$newCode = @"
Public Sub HelloFromPs()
    "MsgBox "Hello From Ps!"
End Sub
"@

$startLine = $codeModule.CountOfLines - 1
$codeModule.InsertLines($startLine, $newCode)
$workbook.Save()
$excel.Quit()

# Clean
[System.Runtime.InteropServices.Marshal]::ReleaseComObject(($excel)) | Out-Null