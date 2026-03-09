# TestWord2PDF.ps1 - Simple standalone Word-to-PDF converter for testing.
# Usage: powershell -ExecutionPolicy Bypass -File TestWord2PDF.ps1 "C:\path\to\file.docx"
param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile
)

$pdfPath = [System.IO.Path]::ChangeExtension($InputFile, '.pdf')
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($InputFile)
$doc.SaveAs([ref]$pdfPath, [ref]17)
$doc.Close([ref]$false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Host "Saved: $pdfPath"
