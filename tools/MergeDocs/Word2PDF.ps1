# Word2PDF.ps1 - Convert Word documents to PDF via COM automation.
# Accepts a file listing input paths and an output directory.
param(
    [Parameter(Mandatory=$true)]
    [string]$InputList,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir
)

$InputFiles = @(Get-Content -LiteralPath $InputList | Where-Object { $_.Trim() -ne '' })

$word = $null
$exitCode = 0

try {
    $word = New-Object -ComObject Word.Application

    foreach ($docFull in $InputFiles) {
        $docName = [System.IO.Path]::GetFileName($docFull)
        $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($docFull) + '.pdf'
        $pdfFull = [System.IO.Path]::Combine($OutputDir, $pdfName)

        try {
            $doc = $word.Documents.Open($docFull)
            $doc.SaveAs([ref]$pdfFull, [ref]17)
            $doc.Close([ref]$false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            $doc = $null

            Write-Output $pdfFull
        }
        catch {
            Write-Error "Failed to convert '$docName': $_"
            $exitCode = 1

            if ($doc -ne $null) {
                try { $doc.Close([ref]$false) } catch {}
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
                $doc = $null
            }
        }
    }
}
catch {
    Write-Error "Failed to start Word: $_"
    $exitCode = 1
}
finally {
    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        $word = $null
    }
    [System.GC]::Collect()
}

exit $exitCode
