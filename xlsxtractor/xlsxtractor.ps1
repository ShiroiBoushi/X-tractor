$rootDirectory = "C:\Users\test\folder_to_analyze" #Change this by your folder
$count = 0
$total = 0 # Number of files that contains the keywords
$keywords = Get-Content .\keywords.txt
Out-File .\result.txt

# Recursively search for XLSX files in the root directory and its subdirectories
Get-ChildItem -Path $rootDirectory -Recurse -Include "*.xlsx" | ForEach-Object {
    $xlsxFilePath = $_.FullName

		# Extract the XLSX file as a zip file
    $zipFilePath = "C:\Users\test\extracted.zip" #zip destination file to change
    $test = Test-Path $($zipFilePath.Replace('.zip', ''))
		
		# Check if folder with the same name exist
    if ($test -eq $True){
    Write-Host "folder" $($zipFilePath.Replace('.zip', ''))"exist"
    exit
		}
		# Extract xslx to destination path
    Copy-Item -Path $xlsxFilePath -Destination $zipFilePath
    Expand-Archive -LiteralPath $zipFilePath -DestinationPath "$($zipFilePath.Replace('.zip', ''))"
    
		# Search for keyword inside SharedStrings.xml
    $sharedStringsFile = "$($zipFilePath.Replace('.zip', ''))\xl\sharedStrings.xml"
		foreach ($keyword in $keywords) {
    if ((Get-Content $sharedStringsFile) -match $keyword) {
        Write-Host "Found $keyword inside $xlsxFilePath"
        Add-Content -Path .\result.txt -Value "Found $keyword inside $xlsxFilePath"
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
