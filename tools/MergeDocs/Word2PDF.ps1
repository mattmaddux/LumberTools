# *********************************************************************
# Script : Word2PDF.ps1
# Based on: https://gist.github.com/TriLogic/faf024344b977f67f468dd10ec570099
# Purpose: Convert a Word document to PDF via COM automation.
#          Accepts one or more input files and an output directory.
# *********************************************************************
param(
    [Parameter(Mandatory=$true)]
    [string]$InputList,

    [Parameter(Mandatory=$true)]
    [string]$OutputDir
)

# Read file paths from the list file (one per line)
$InputFiles = @(Get-Content -LiteralPath $InputList | Where-Object { $_.Trim() -ne '' })

$appWord = $null
$docWord = $null
$exitCode = 0

try {
    $appWord = New-Object -ComObject Word.Application
    # Word starts hidden by default via COM; skip setting Visible to avoid
    # TYPE_E_CANTLOADLIBRARY on machines with broken Office interop assemblies.

    foreach ($docFull in $InputFiles) {
        $docName = [System.IO.Path]::GetFileName($docFull)
        $pdfName = [System.IO.Path]::GetFileNameWithoutExtension($docFull) + '.pdf'
        $pdfFull = [System.IO.Path]::Combine($OutputDir, $pdfName)

        try {
            $docs = $appWord.GetType().InvokeMember("Documents", "GetProperty", $null, $appWord, $null)
            $docWord = $docs.GetType().InvokeMember("Open", "InvokeMethod", $null, $docs, @($docFull))
            $docWord.GetType().InvokeMember("ExportAsFixedFormat", "InvokeMethod", $null, $docWord, @($pdfFull, 17))
            $docWord.GetType().InvokeMember("Close", "InvokeMethod", $null, $docWord, @(0))

            while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($docWord)) {}
            $docWord = $null
            [System.GC]::Collect()

            Write-Output $pdfFull
        }
        catch {
            Write-Error "Failed to convert '$docName': $_"
            $exitCode = 1

            if ($docWord -ne $null) {
                try { $docWord.GetType().InvokeMember("Close", "InvokeMethod", $null, $docWord, @(0)) } catch {}
                while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($docWord)) {}
                $docWord = $null
                [System.GC]::Collect()
            }
        }
    }
}
catch {
    Write-Error "Failed to start Word: $_"
    $exitCode = 1
}
finally {
    if ($appWord -ne $null) {
        $appWord.GetType().InvokeMember("Quit", "InvokeMethod", $null, $appWord, $null)
        while ([System.Runtime.InteropServices.Marshal]::ReleaseComObject($appWord)) {}
        $appWord = $null
    }
    [System.GC]::Collect()
}

exit $exitCode
