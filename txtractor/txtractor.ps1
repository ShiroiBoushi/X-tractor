$rootDirectory = "C:\Users\test\folder_to_analyze" #Change this by the folder you want to analyze
$count = 0
$total = 0
$keywords = Get-Content .\keywords.txt

# Recursively search for txt files in the root directory and its subdirectories
Get-ChildItem -Path $rootDirectory -Recurse -Include "*.txt" | ForEach-Object {
    $txtFilePath = $_.FullName
    # Search for keywords inside txt file
		foreach ($keyword in $keywords) {
    if ((Get-Content $txtFilePath) -match $keyword) {
        Write-Host "Found MDP inside $txtFilePath"
        $count++
    }
    if ($count -ge 1){
    $total = $total + 1}
    $count = 0
}
}
Write-Host $total "fichier(s) trouvé(s)"
