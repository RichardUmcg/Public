<#
.SYNOPSIS
  The script migrates unstructured files and folders into a known folder. That folder
  could then be used as a known folder that can be migrated automatically.
.DESCRIPTION
  The script must run in a users context. It scans the root of a given (home) folder. It then
  'substracts' the Known Folders (ie. MyDocuments, Pictures) and other that must not be migrated into a
  standardized folder. What remains is all the other 'unknown' files and folders that will be moved into 
  a predefined known folder that can be migrated automatically.
.INPUTS
  None
.OUTPUTS
  Log file stored in the root of the migration folder: 
  %HOMEDRIVE%%HOMEPATH%\nondefaultfolders.migrated\_migration.log
.NOTES
  Version:        1.0
  Author:         Richard Wijngaard
  Creation Date:  05/10/2022
  Purpose/Change: Initial script development
#>

#region variables

$SourceFolder = $env:HOMEDRIVE + $Env:HOMEPATH
$MigrationFolderName = "nondefaultfolders.migrated"
$DestinationFolder = $SourceFolder + "\" + $MigrationFolderName
$logfile = $DestinationFolder + "\_migration.log"

#region To be copied
$foldersToMigrate = [System.Collections.ArrayList] (Get-ChildItem $SourceFolder -Directory -Name)
$FilesInRoot = [System.Collections.ArrayList] (Get-ChildItem $SourceFolder -File -Name)
#endregion

$Knownfolders = @(
        [environment]::getfolderpath(“myDocuments”), #must be first!
        [environment]::getfolderpath(“myPictures”), 
        [environment]::getfolderpath(“MyMusic”), 
        [environment]::getfolderpath(“MyVideos”), 
        [environment]::getfolderpath(“Desktop”)
    )
    #https://gist.github.com/DamianSuess/c143ed869e02e002d252056656aeb9bf#all-enums

$SubfoldersToExcludeFromMigration = [System.Collections.ArrayList] ('Downloads', 'AppData', $MigrationFolderName) 

#endregion

#region Functions area

<#
.SYNOPSIS
Adds a log entry to a logfile.

.DESCRIPTION
Adds a log entry to a logfile.cls

.PARAMETER Level
Type of log entry: "INFO","WARN","ERROR","FATAL","DEBUG". Optional. When blank then "INFO"

.PARAMETER Message
Description of the log entry

.PARAMETER logfile
Path to the logfile

.EXAMPLE
Write-Log -Message 'Text to log' -Level 'WARN'
.EXAMPLE
Write-Log -Message 'Text to log'
.EXAMPLE
Write-Log 'Text to log'

.NOTES
General notes
#>
Function Write-Log {
    [CmdletBinding()]
    Param(

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO"
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

<#
.SYNOPSIS
Write what will (not) be migrated to the logging.

.DESCRIPTION
Write what will (not) be migrated to the logging.
#>
Function Write-Workload  {
    $tekst = "Deze mappen worden niet verplaatst:`n" + (Format-ArrayList -ListToFormat $SubfoldersToExcludeFromMigration)
    Write-Log -Message $tekst
    $tekst = "Deze mappen worden verplaatst:`n" + (Format-ArrayList -ListToFormat $foldersToMigrate)
    Write-Log -Message $tekst
    $tekst = "Deze bestanden worden verplaatst:`n" + (Format-ArrayList -ListToFormat $FilesInRoot)
    Write-Log -Message $tekst
}

<#
.SYNOPSIS
Do some housekeeping before doing the job.

.DESCRIPTION
Do some housekeeping before doing the job.

#>
Function Initialize-Script
{
    If (!(Test-Path -Path $logfile)) {New-Item -Path $logfile -Force}
    Write-Log -Message "*** Start Migration ***"
}

#region string manupulation functions
<#
.SYNOPSIS
Makes a sweet formated list in a string.

.DESCRIPTION
Makes a sweet formated list in a string.

.PARAMETER ListToFormat
ArrayList to be formated into a string.

.EXAMPLE
Format-ArrayList -ListToFormat $ArrayListToFormat

#>
Function Format-ArrayList {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$True)]
    [System.Collections.ArrayList]
    $ListToFormat)

    $result = ""
    foreach ($formatItem in $ListToFormat) {
        $result += "`t$formatItem`n"
    }
    return $result
}

<#
.SYNOPSIS
To set the full path to a relative name of a file or folder.

.DESCRIPTION
To set the full path to a relative name of a file or folder.

.PARAMETER TheList
The ArrayList with items to provide a full path

.EXAMPLE
Set-FullPath -TheList $ArrayList with items
#>
Function Set-FullPath{
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$True)]
    [System.Collections.ArrayList]
    $TheList)

    $resultList = [System.Collections.ArrayList]@()

    foreach ($listItem in $TheList) {
        if (!($listItem.StartsWith($SourceFolder))) {
            $listItem = $SourceFolder + "\" + $listItem
        }
        [void] $resultList.Add($listItem)
    }
    return $resultList
}
#endregion

#region Select the files and directories to move
<#
.SYNOPSIS
Gets an arraylist of subfoldernames where certain foldersToMigrate ($SubfoldersToExcludeFromMigration) are excluded.

.DESCRIPTION
Gets an arraylist of subfoldernames where certain foldersToMigrate ($SubfoldersToExcludeFromMigration) are excluded.
#>
Function Set-IncludedfoldersToMigrate {

    # Excluding some local sync-foldersToMigrate from cloud services in the users home drive.
    # Synchronized foldersToMigrate shouldn't be included, because they can be sync'ed again
    # and forces to be downloaded before being able to move.
    Set-KnownfoldersToMigrate
    Set-ExcludeOneDriveForBusiness
    Set-ExcludeDropBox
    Set-ExcludeiCloudDrive

    #Google does not store sync-foldersToMigrate in users home drive.

    #Find foldersToMigrate to be copied
    foreach ($folder in $SubfoldersToExcludeFromMigration) {

        if ($foldersToMigrate.Contains($folder)) {
            
            $foldersToMigrate.Remove($folder)
        }
    }

    #Set the full path for the relative names in the ArrayLists globaly for further processing.
    Set-Variable -Name "SubfoldersToExcludeFromMigration" -Value (Set-FullPath -TheList $SubfoldersToExcludeFromMigration) -Scope Global
    Set-Variable -Name "foldersToMigrate" -Value (Set-FullPath -TheList $foldersToMigrate) -Scope Global
    Set-Variable -Name "FilesInRoot" -Value (Set-FullPath -TheList $FilesInRoot) -Scope Global

    Write-Workload
}

<#
.SYNOPSIS
Add the known foldersToMigrate to be excluded from the list of to migrate foldersToMigrate.

.DESCRIPTION
Add the known foldersToMigrate to be excluded from the list of to migrate foldersToMigrate.
The idea is that with the use of reading the known path folder from the system makes sure
the right name is used. This avoids problems due to the users language. foldersToMigrate with names
that are known folder names in another language will be treated as a regular folder that
must be moved.
Known foldersToMigrate nested in another known folder are also skipped, because they haven't have a
folder in the root of the users homedrive. (And will be migrated otherwise.)

#>
Function Set-KnownfoldersToMigrate {
    
    foreach ($known in $KnownfoldersToMigrate) {
        if ((Get-IsNested($known)) -eq $false) {
            $tmp = $known.Split("\")
            if ([string]::IsNullOrEmpty($tmp.Count - 1) -eq $false) {
                [void]$SubfoldersToExcludeFromMigration.Add($tmp[$tmp.Count - 1])
            }
        }
    }
}

<#
.SYNOPSIS
Test if folder is subfolder of the MyDocuments folder.

.DESCRIPTION
Test if folder is subfolder of the MyDocuments folder.

.PARAMETER Path
Mandatory path to examine if it is a subfolder of MyDocuments

#>
function Get-IsNested {
    Param(
    [parameter(Mandatory=$true)]
    [String] $Path
    )

    if ($Path -eq $Knownfolders[0]) {return $false} #Path = MyDocuments
    if ($Path.StartsWith($Knownfolders[0])) {return $true}
    return $false
}

<#
.SYNOPSIS
Exclude OneDrive for Business

.DESCRIPTION
Excluding OneDrive for Business (Multiple accounts too), because it's a sync'd folder with the cloud in the users homedrive.
#>
Function Set-ExcludeOneDriveForBusiness {
    $OD4B = $foldersToMigrate.where{$_ -match 'OneDrive -'}

    foreach ($odFolder in $OD4B) {
        $SubfoldersToExcludeFromMigration.Add($odFolder)
    }
}

<#
.SYNOPSIS
Exclude Dropbox

.DESCRIPTION
Excluding DropBox.
#>
Function Set-ExcludeDropBox {
    if ($foldersToMigrate.Contains("Dropbox")) {
            
        $SubfoldersToExcludeFromMigration.Add("Dropbox")
    }
}

<#
.SYNOPSIS
Exclude iCloud Drive

.DESCRIPTION
Excluding iCloud Drive.
#>
Function Set-ExcludeiCloudDrive {
    if ($foldersToMigrate.Contains("iCloud Drive")) {
            
        $SubfoldersToExcludeFromMigration.Add("iCloud Drive")
    }
}
#endregion

#region Moving data functions
<#
.SYNOPSIS
Move the selected files to the destination location.

.DESCRIPTION
Move the selected files to the destination location.
#>
function Move-FilesInRoot {

    $countFiles =0
    Write-Log -Message $FilesInRoot.Count + " bestanden te verplaatsen..."

    foreach ($fileItem in $FilesInRoot) {
        try {
            Move-Item -Path $fileItem -Destination $DestinationFolder -ErrorAction Stop -WhatIf
            $countFiles += 1  
        }
        catch {
            Write-Log -Message "Verplaatsen van $fileItem is mislukt" -Level "ERROR"
        }
    }
    Write-Log -Message "Er zijn $countFiles bestanden succesvol verplaatst. Bij " 
        + ($FilesInRoot.Count - $countFiles) + " bestanden is er iets misgegaan."
}

<#
.SYNOPSIS
Move the selected folders to the destination location.

.DESCRIPTION
Move the selected folders to the destination location.
#>
function Move-SelectedFolders {
    $countFolders = 
    Write-Log -Message $foldersToMigrate.Count "folders te verplaatsen..."

    foreach($folderItem in $foldersToMigrate) {
        try {
            Move-Item -Path $folderItem -Destination $DestinationFolder -ErrorAction Stop -WhatIf
            $countFolders += 1
        }
        catch {
            Write-Log -Message "Verplaatsen van $folderItem is mislukt." -Level "ERROR"
        }
        
    }
    Write-Log -Message "Er zijn $countFolders folders succesvol verplaatst. Bij "
        + ($foldersToMigrate.Count - $countFolders) + "folders is er iets misggegaan."

}

<#
.SYNOPSIS
Do the move from the included files and directories.

.DESCRIPTION
Do the move from the included files and directories.

#>
function Move-FilesAndFolders {
    Write-Log -Message "Start migratie van bestanden in de root van $SourceFolder."
    Move-FilesInRoot
    Write-Log -Message "Migratie van bestanden voltooid."

    Write-Log -Message "Start migratie van de folders."
    Move-SelectedFolders
    Write-Log -Message "Migratie van folders voltooid."
}
#endregion

#endregion

Initialize-Script
Set-IncludedfoldersToMigrate
#Move-FilesAndFolders  #MOve is not tested. It also use -WhatIf in the move commands. Should be removed for testing.

Write-Log -Message "*** De migratie van bestanden en folders naar $SourceFolder is voltooid. ***"
