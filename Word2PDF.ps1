# *********************************************************************
# Script : Word2PDF.ps1
# Author : Andrew F. Friedl
# Purpose: Convert Word type documents to PDF (*.txt, *.rtf, *.doc?)
# Created: 2017.06.22
# Source : https://gist.github.com/TriLogic/faf024344b977f67f468dd10ec570099
# *********************************************************************#
#Clear-Host

# Path definitions
$inpDocPath = '${WordSrcPath}'
$outDocPath = '${WordTarPath}'
$arcDocPath = '${WordArcPath}'
$errDocPath = '${WordErrPath}'
$appWord = $null
$docWord = $null

# Log folder locations
Write-Output "inpDocPath = $inpDocPath"
Write-Output "outDocPath = $outDocPath"
Write-Output "arcDocPath = $arcDocPath"
Write-Output "errDocPath = $errDocPath"

# Convert all Word documents
dir -Path ([System.IO.Path]::Combine($inpDocPath,'*.*')) -Include *.rtf,*.txt,*.doc*,*.odt | ForEach-Object {

	if ($appWord -eq $null) {
			$appWord = New-Object -ComObject Word.Application
			$appWord.Visible = $false
		}

	$docFull = $_.FullName
	$docPath = [System.IO.Path]::GetDirectoryName($docFull)
	$docName = [System.IO.Path]::GetFileName($docFull)
	$pdfType = 'Microsoft.Office.Interop.Word.WdSaveFormat' -as [type]
	$pdfPath = [System.IO.Path]::GetDirectoryName($docFull)
	$pdfName = ($outXlsPref + [System.IO.Path]::GetFileNameWithoutExtension($docFull))
	$pdfFull = [System.IO.Path]::Combine($outDocPath, $pdfName + '.pdf')

Write-Output '--------------------------------------------------------------------------------'
Write-Output "docFull = $docFull"
Write-Output "docPath = $docPath"
Write-Output "docName = $docName"
Write-Output "pdfPath = $pdfPath"
Write-Output "pdfName = $pdfName"
Write-Output "pdfFull = $pdfFull"

	Try {

		# Open, Save as a PDF in the final location and close the original
		$docWord = $appWord.Documents.Open($docFull)
		$docWord.ExportAsFixedFormat($pdfFull, 17)
		$docWord.Close([ref]$false)

		# Release the document
		while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($docWord)){}
		[System.GC]::Collect()
		$docWord = $null

		# Log conversion success
		Write-Output ('Converted: ' + $docName)

		# Move the original file to the archive location
		Push-Location $docPath
		Move-Item $docName $arcDocPath
		Pop-Location

	} Catch {

		# Log conversion failure
		Write-Output ('Conversion failed: ' + $docName)

		If ($docWord -ne $null) {
			$docWord.Close([ref]$false)
			while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($docWord)){}
			$docWord = $null
			[System.GC]::Collect()
			}

		# Move the original file to the error location
		Push-Location $docPath
		Move-Item $docName $errDocPath
		Pop-Location

	} Finally {

		# Free all references to the Word document object
		If ($docWord -ne $null) {
			while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($docWord)){}
		}

		$docWord = $null
		$pdfType = $null
	}

	# Force CLR to Garbage Collect
	[System.GC]::Collect()

}

# Close Word if Opened
if ($appWord -ne $null) {
	$appWord.Quit()
	while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($appWord)) {}
	$appWord = $null;
	}

# force garbage collection
[System.GC]::Collect()

# Exit the script
Exit 0
