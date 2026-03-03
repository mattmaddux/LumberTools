@echo off
if "%~1"=="" (
    echo Drag a Word file onto this bat to convert it to PDF.
    pause
    exit /b 1
)
powershell.exe -ExecutionPolicy Bypass -Command ^
    "$word = New-Object -ComObject Word.Application; $word.Visible = $false; $doc = $word.Documents.Open('%~1'); $pdf = [System.IO.Path]::ChangeExtension('%~1', '.pdf'); $doc.ExportAsFixedFormat($pdf, 17); $doc.Close([ref]$false); $word.Quit(); Write-Host ('Saved: ' + $pdf)"
pause
