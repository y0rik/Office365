[CmdletBinding()]
Param (
	[Parameter(Position = 0, Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[String[]]$PstFiles,
  [Parameter(Position = 1, Mandatory = $true)]
	[ValidateNotNullOrEmpty()]
	[String]$RootFolder,
  [Parameter(Position = 2, Mandatory = $false)]
	[switch]$CheckLongPathsEnabled = $true,
  [Parameter(Position = 3, Mandatory = $false)]
  [switch]$ShowProgressBar = $true
)

#FUNCTIONS

#function builds and returns full file path for the message to be exported
function fGet-FilePath ([string]$_Date, [string]$_Subject, [string]$_rootPath) {

	#check if the subject is empty
  if ($_Subject -eq $null -or $_Subject.length -eq 0) {
		#put NoSubject name for the email
		[string]$_FileName = "NoSubject"
  }
  else {
		#re-construct the subject to exclude some chars
    [string]$_FileName = ""
    for ($_i=0; $_i -lt $_Subject.length; $_i++) {
      if (($([int64][char]"$($_Subject[$_i])") -gt 31) -and ($([int64][char]"$($_Subject[$_i])") -lt 127)) {
        $_FileName += $_Subject[$_i]
      }
    }

		#replace some special chars
    $_FileName = $_FileName.Replace('\\','_')
    $_FileName = $_FileName.Replace('/','_')
    $_FileName = $_FileName.Replace(':','_')
    $_FileName = $_FileName.Replace('*','_')
    $_FileName = $_FileName.Replace('?','_')
    $_FileName = $_FileName.Replace('\','_')
    $_FileName = $_FileName.Replace('<','_')
    $_FileName = $_FileName.Replace('>','_')
    $_FileName = $_FileName.Replace('|','_')
    $_FileName = $_FileName.Replace('"','_')
    $_FileName = $_FileName.Replace("'",'_')

    #build first file path
    $_FilePath = Join-Path -Path "$_rootPath" -ChildPath "$_Date-$_FileName"
		#if there is path limit - cut it
    if ($LimitPathLength -and $_FilePath.length -gt 250) {$_FilePath = $_FilePath.Substring(0,250)}
    $_FilePath = "$_FilePath.msg"

    $_i = 0

    do {
      #check file
      $_pathTest = Test-Path -Path "$_FilePath"
      #if the file exists - search for a proper index
      if ($_pathTest) {
        $_i ++
        $_FilePath = Join-Path -Path "$_rootPath" -ChildPath "$_Date-$_FileName"
        if ($LimitPathLength -and $_FilePath.length -gt 250) {$_FilePath = $_FilePath.Substring(0,250)}
        $_FilePath = "$_FilePath($_i).msg"
      }
    } while ($_pathTest)

		#return the path
    return $_FilePath
  }
}

#function builds PST file folder tree
function fGet-SubFolders ($_folderIDs) {
  $_foldersToReturn = @()
  foreach ($_f in $_folderIDs) {
    $_subFolders = $folders | ?{$_.ParentId -eq $_f}
    foreach ($_sf in $_subFolders) {
      $foldersStructure.Add($($_sf.ID),"$($foldersStructure.Item($_sf.ParentId))/$($_sf.DisplayName)")
      $_foldersToReturn += $($_sf.ID)
    }
  }

  return $_foldersToReturn
}

#function checks if long paths are enabled
function fCheck-LongPathSetting () {
  return $((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name LongPathsEnabled).LongPathsEnabled)
}

#function enables long paths
function fChange-LongPathSetting () {
  $_registryPath = "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem"

  #check if powershell is running as Admin
  #get the ID and security principal of the current user account
  $_myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
  $_myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($_myWindowsID);

  #get the security principal for the administrator role
  $_adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

  #check if we are currently running as an administrator
  if ($_myWindowsPrincipal.IsInRole($_adminRole))
  {
    #test/create the path
    if (-not (Test-Path -Path "$_registryPath")) {New-Item -Path "$_registryPath" -Force}
    #change the value
    Set-ItemProperty -Path "$_registryPath" -Name 'LongPathsEnabled' -Value 1
  }
  else {
		#run a new powershell process as admin to modify the setting
    $_newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $_newProcess.Arguments = "if (-not (Test-Path -Path '$_registryPath')) {New-Item -Path '$_registryPath' -Force}; Set-ItemProperty -Path '$_registryPath' -Name 'LongPathsEnabled' -Value 1"
    $_newProcess.Verb = "runas"
    $_newProcess.WindowStyle = "hidden"
    [System.Diagnostics.Process]::Start($_newProcess)
  }

	#verify that the setting is successfully applied
  return fCheck-LongPathSetting
}

#BODY
#paths
$ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$ScriptName = $MyInvocation.MyCommand.Name

#stats
$ProcessedItems = @()

#check root folder parameter
if ($RootFolder.length -eq 0) {
  Write-Warning "No -RootFolder parameter is specified. Script root folder will be used: $ScriptPath"
  $RootFolder = $ScriptPath
}

Write-Host "Checking long paths setting" -ForegroundColor DarkGray
#check long paths
if ($CheckLongPathsEnabled) {
  if (-not $(fCheck-LongPathSetting)) {
    Write-Warning "Long paths are not enabled on this computer. This means .msg files will be exported with shorten names limited to 260 characters of full path"
    #ask if it's okay to proceed with short paths
    do {
      $return = Read-Host -Prompt "Do you want to enable long paths?(yes/no)"
      if (($return -notlike "yes") -and ($return -notlike "no")) {
        Write-Host -ForegroundColor "Red" "Please, type 'yes' or 'no'"
      }
    } until (($return -like "yes") -or ($return -like "no"))

    if ($return -like "yes") {
      $returnSuccess = fChange-LongPathSetting
      if (-not $returnSuccess) {
        Write-Error "Long paths setting haven't been changed"
        Exit
      }
      else {
        Write-Host "Long paths setting have been changed. Please restart your computer for changes to take effect" -ForegroundColor DarkGray
        Exit
        #$LimitPathLength = $false
      }
    }
    elseif ($return -like "no") {
      Write-Warning "Continue with 260 characters limitation"
      $LimitPathLength = $true
    }
  }
  else {
    Write-Host "Long paths are already enabled" -ForegroundColor DarkGray
    $LimitPathLength = $false
  }
}
else {
  Write-Warning "Long paths check skipped. This means .msg files will be exported with shorten names limited to 260 characters of full path"
  $LimitPathLength = $true
}


#use the assembly
$libPath = "$ScriptPath\Independentsoft.Pst.dll"
if (Test-Path -Path "$libPath") {
  Write-Host "Loading library: '$libPath'" -ForegroundColor DarkGray
  [Reflection.Assembly]::LoadFile("$libPath")
}
else {
  Write-Error "No mandatory file found: $libPath"
  Exit
}

foreach ($PstFilePath in $PstFiles) {

  Write-Host "Processing the PST file: '$PstFilePath'" -ForegroundColor DarkGray

  #import PST file
  try {$PSTFile = New-Object Independentsoft.Pst.PstFile("$PstFilePath")}
  catch {
    Write-Error "Cannot open PST file: '$PstFilePath'"
    Exit
  }

  #get folders recursively
  $folders = $PSTFile.MailboxRoot.GetFolders($true)

  #ID-FullPath dictionary
  $foldersStructure = New-Object 'System.Collections.Generic.Dictionary[string,string]'

  #add the root folder
  $foldersStructure.Add($($PSTFile.MailboxRoot.Id),"/$($PSTFile.MailboxRoot.DisplayName)")

  #build path for each folder
  $returnedFolders = @($PSTFile.MailboxRoot.Id)
  while ($returnedFolders.count -ne 0) {
    $returnedFolders = fGet-SubFolders $returnedFolders
  }

  #get root folder for the PST file
  $pstRootFolder = $folderPath = Join-Path -Path "$RootFolder" -ChildPath "$($PstFilePath.Split('\')[-1])"

  #if path limitation is in place - check root length
  if ($LimitPathLength -and $($pstRootFolder.length) -ge 180) {
    Write-Error "PST file root folder path is longer than 180 symbols, cannot proceed: $pstRootFolder"
    Exit
  }

  $PSTStatsObject = New-Object PSObject -Property @{
      PSTFile = $PstFilePath
      PSTExportRootFolder = $pstRootFolder
  }
  $ProcessedItems += $PSTStatsObject

  #check if the folder exists
  $return = ''
  if (Test-Path -Path "$pstRootFolder" -PathType 'Container') {
    Write-Warning "PST file root folder exists and should be deleted in order to proceed: '$pstRootFolder'"
    #ask if it's okay to proceed with short paths
    do {
      $return = Read-Host -Prompt "Do you want to delete it?(yes/no)"
      if (($return -notlike "yes") -and ($return -notlike "no")) {
        Write-Host -ForegroundColor "Red" "Please, type 'yes' or 'no'"
      }
    } until (($return -like "yes") -or ($return -like "no"))
  }
  elseif (Test-Path -Path "$pstRootFolder" -PathType 'Leaf') {
    Write-Error "PST root folder cannot be created due to existing file with the same name: '$pstRootFolder'. Try choosing another -RootFolder parameter or deleting existing file: '$pstRootFolder'"
    Exit
  }

  if ($return -like 'yes') {
    Write-Host "Removing the folder: '$pstRootFolder'" -ForegroundColor DarkGray
    try {Remove-Item -Path "$pstRootFolder" -Recurse -Force}
    catch {Write-Error "Cannot fully delete the folder: '$pstRootFolder'"; Exit}
  }
  elseif ($return -like 'no') {
    Write-Error "Cannot continue without the folder deletion: '$pstRootFolder'"
    Exit
  }

  $fCount = 0
  $fTotal = $folders.count
  #export items
  foreach ($f in $folders) {
    $folderPath = Join-Path -Path "$pstRootFolder" -ChildPath "$($foldersStructure.Item($f.Id))"
    $fCount ++
    $folderItemsCount = $f.ChildrenCount

    Write-Host "Processing the folder: '$folderPath'" -ForegroundColor DarkGray
    #check if folder exists and create if needed
    if (-not (Test-Path -Path "$folderPath")) {
      Write-Host "->Creating the folder: '$folderPath'" -ForegroundColor DarkGray
      try {New-Item -Path "$folderPath" -ItemType 'Directory' >> $null}
      catch {
        Write-Warning "->Error creating: '$folderPath'"
        Write-Error "Cannot continue without the folder: '$folderPath'"
        Exit
      }
    }

    #export all messages from the folder
    for ($i = 0; $i -le $folderItemsCount; $i += 100) {

      #status update
      if ($folderItemsCount -ne 0 -and $ShowProgressBar) {
        Write-Progress -Activity "Processing PST file: '$PstFilePath'; Folder #$fCount of total $fTotal. Current: '$folderPath'" -Status "Messages processed $i of total $folderItemsCount" -PercentComplete $(($i/$folderItemsCount)*100)
      }

      $messages = $f.GetItems($i, $i + 100)
      #processing messages
      foreach ($m in $messages) {
        #build the path
        $msgPath = fGet-FilePath -_Date $($m.CreationTime.ToString('yyyyMMdd-HHmmss')) -_Subject $($m.Subject) -_rootPath $folderPath
        try {
          #saving th file
          $m.Save("$msgPath")
          Write-Host "->File processed: '$msgPath'" -ForegroundColor Green
        }
        catch {Write-Warning "->Error processing: '$msgPath'"}

      }
    }

  }
}

Write-Host "The script is completed. Please find results below" -ForegroundColor DarkGray
$ProcessedItems | ft -a
