$rootDirectory = "C:\Users\test\folder_to_analyze" #Change this by the folder you want to analyze
$count = 0
$total = 0 # Number of files that contains the keywords
$keywords = Get-Content .\keywords.txt
Out-File .\result.txt

# Recursively search for DOCX files in the root directory and its subdirectories
Get-ChildItem -Path $rootDirectory -Recurse -Include "*.docx" | ForEach-Object {
    $docxFilePath = $_.FullName

		# Extract the DOCX file as a zip file
    $zipFilePath = "C:\Users\test\extracted.zip" #zip destination folder to change
    $test = Test-Path $($zipFilePath.Replace('.zip', ''))
		
		# Check if folder with the same name exist
    if ($test -eq $True){
    Write-Host "folder" $($zipFilePath.Replace('.zip', ''))"exist"
    exit
		}
		# Extract docx to destination path
    Copy-Item -Path $docxFilePath -Destination $zipFilePath
    Expand-Archive -LiteralPath $zipFilePath -DestinationPath "$($zipFilePath.Replace('.zip', ''))"
    
		# Search for keywords inside SharedStrings.xml
    $sharedStringsFile = "$($zipFilePath.Replace('.zip', ''))\word\document.xml"
		foreach ($keyword in $keywords) {
    if ((Get-Content $sharedStringsFile) -match $keyword) {
        Write-Host "Found $keyword inside $docxFilePath"
				Add-Content -Path .\result.txt -Value "Found $keyword inside $docxFilePath"
        $count++
    }
    }
    # Delete the extracted files and zip file
    Remove-Item "$($zipFilePath.Replace('.zip', ''))" -Recurse -Force
    Remove-Item $zipFilePath
    if ($count -ge 1){
    $total++
		}
    $count = 0
}
Write-Host "$total fichiers trouvés"
Add-Content -Path .\result.txt -Value "$total fichiers trouvés"
# Convertir en .csv
Get-ChildItem result.txt | Rename-Item -NewName {$_.Name -replace '\.txt$','.csv'}
