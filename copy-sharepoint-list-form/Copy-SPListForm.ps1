# Information about the list that has the Power App form associated with it
$global:srcListId = "<<Id (guid) of the list that has the customized form>>"
$global:srcListName = "<<Name of the list that has the customized form>>"
$global:srcListUrl = "<<Web address of the list that has the customized form>>"

# Information about the list that you want to associate the  Power App form with
$global:destListId = "<<Id (guid) of the list to move the customized form to>>"
$global:destListName = "<<Name of the list to move the customized form to>>"
$global:destListUrl = "<<Web address of the list to move the customized form to>>"

# App package from the source list
$global:srcAppPackage = "<<Path of the downloaded zip file>>"

# New app package set for destination list 
$global:destAppPackage = "<<Path of the new zip file>>"

# ************ DO NOT MODIFY ANYTHING BELOW THIS LINE ************

function Write-Log ([string]$Message, [string]$LogFilePath){
  # Get the current date and time in the computer's timezone
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"

  # Prefix the message with the timestamp
  $logMessage = "[$timestamp] $Message"

  #Write-Host $logMessage

  # Append the message to the log file
  Add-Content -Path $LogFilePath -Value $logMessage
}

function Find-MSAppFile([string]$folder) {
  Get-ChildItem -Path $folder -Filter *.msapp -Recurse -File | ForEach-Object {    
    return $_.FullName
  }
}

function Edit-Files($FolderPath, $LogFilePath) {
  Get-ChildItem -Path $FolderPath -Filter *.json -Recurse -File | ForEach-Object {
    Write-Log -Message "Modifying file $($_.FullName)" -LogFilePath $LogFilePath
    Edit-File -FilePath $_.FullName
  }

  Get-ChildItem -Path $FolderPath -Filter *.yaml -Recurse -File | ForEach-Object {
    Write-Log -Message "Modifying file $($_.FullName)" -LogFilePath $LogFilePath
    Edit-File -FilePath $_.FullName
  }
}

function Edit-File([string]$filePath) {
  # Read the file content
  $fileContent = (Get-Content -Path $filePath -Raw)

  # Replace the old text with the new text
  $fileContent = $fileContent -replace $global:srcListUrl, $global:destListUrl
  $fileContent = $fileContent -replace $global:srcListName, $global:destListName
  $fileContent = $fileContent -replace $global:srcListId, $global:destListId

  # Write out the new file content
  Set-Content -Path $filePath -Value $fileContent
}

### TODO: Check if the Power Platform CLI is installed

# Get the absolute path in case only the relative path was passed in
$srcAppPackage = (Get-ChildItem -Path $srcAppPackage).FullName

# Gets the parent folder of the app package
$srcAppLocation = Split-Path -Parent $srcAppPackage

$logFile = -join([System.IO.Path]::GetFileNameWithoutExtension($destAppPackage), ".log")

$workingFolder = -join($srcAppLocation, "\", [guid]::NewGuid().ToString())

# Folder where the app package will be extracted
$appPackageFolder = -join($workingFolder, "\zip")
$appPackageFolderZipContents = -join($appPackageFolder, "\*")

# Folder where the .msapp file will be extracted
$msAppFolder = -join($workingFolder, "\msapp")

# Extract the app package
Write-Log -Message "Extracting $srcAppPackage to $appPackageFolder" -LogFilePath $logFile
Expand-Archive -Path $srcAppPackage -DestinationPath $appPackageFolder -Force

# Unpack msapp file
$msAppFile = Find-MSAppFile -folder $appPackageFolder
Write-Log -Message "Unpacking $msAppFile to $msAppFolder" -LogFilePath $logFile
pac canvas unpack --msapp $msAppFile --sources $msAppFolder

Edit-Files -FolderPath $workingFolder -LogFilePath $logFile


Write-Log -Message "Packing $msAppFile" -LogFilePath $logFile
pac canvas pack --sources $msAppFolder --msapp $msAppFile

# Zip up the app package
Write-Log -Message "Zipping up $appPackageFolder to $destAppPackage" -LogFilePath $logFile
Compress-Archive -Path $appPackageFolderZipContents -DestinationPath $destAppPackage -Force

Write-Log -Message "Deleting $workingFolder" -LogFilePath $logFile
Remove-Item -Path $workingFolder -Recurse -Force