# X-tractor
Powershell scripts that searches for strings inside Office (docx, xlsx) and txt files recursively and list paths and count the number of files containing these keywords.
# Tools
-xlsxtractor

-docxtractor

-txtractor
# Usage
Edit the powershell script to choose the source folder and where to extract the Office files (See below why).

Edit the **keywords.txt** file and seperate the keywords with line break like this :
```
password
credential
secret
...
```
Run the script from a folder and it will search the folder and subfolders.
```
xlsxtractor.ps1
```
# How it works
Open XML formatted documents are ZIP files containing XML files that contains text you typed on the Word page or Excel sheet.
The scripts will make a copy of the files with .xlsx extension to a destination folder outside the folder you search (in order not to alter the original content) and rename it with .zip extension. Then it will extract the archive and search for the keywords inside **xl\SharedStrings.xml** (an error will occur if the file is empty because no SharedStrings.xml but while not stop the script from running). Once the copy of the file (in the destination folder) has been analyzed, it is deleted to continue with the other files.

Note: If a file has the same name as a folder (test.xlsx and a folder is named test) the script will stop before extract and delete (in the case the destination folder is inside the folder that is analysed to not delete a folder).

# Use cases
If you have to audit a file server to find cleartext passwords inside it and migrate them to password manager.

If you want to find a string inside many files and don't remember in which file you wrote it.

# To do
PPTX files

Other file extensions

# Warning
Do not use these scripts on files that you are not authorized to access.
