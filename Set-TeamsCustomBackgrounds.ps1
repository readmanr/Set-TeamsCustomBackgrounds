function Initialise-EventLogging {
    param (
        [Parameter(Mandatory=$false)] [string]$logName = "Application",
        [Parameter(Mandatory=$false)] [string]$source
    )

    # Create the source if it does not exist
    if (![System.Diagnostics.EventLog]::SourceExists($source)) {
        $message = "Initialise-EventLogging @ "+(Get-Date)+": Creating LogSource for EventLog..."
        Write-Verbose $message
        [System.Diagnostics.EventLog]::CreateEventSource($source, $logName)
    } else {
        $message = "Initialise-EventLogging @ "+(Get-Date)+": LogSource exists already."
        Write-Verbose $message
    }
}

function Log-Event {
    param (
        [Parameter(Mandatory=$false)] [string]$logName = "Application",
        [Parameter(Mandatory=$false)] [string]$source = "Intune-PoSh-SPE-WIN10-Baseline-Devices-Teams-CustomBackgrounds",
        [Parameter(Mandatory=$false)] [string]$entryType = "Information",
        [Parameter(Mandatory=$false)] [int]$eventId = 1001,
        [Parameter(Mandatory=$true)] [string]$message
    )

    Write-EventLog -LogName $logName -Source $source -EntryType $entryType -EventId $eventId -Message $message
}

Function Initialise-TeamsLocalUploadFolder {
    param (
        [Parameter(Mandatory=$false)] [boolean]$IncludeNewTeams
    )

    $TeamsBackgroundBasePath = $env:APPDATA+"\Microsoft\Teams\Backgrounds\"
    $TeamsBackgroundUploadPath = $TeamsBackgroundBasePath+"\Uploads\"

    if (!(Test-Path $TeamsBackgroundUploadPath)) {
        $message = "Initialise-TeamsLocalUploadFolder @ "+(Get-Date)+": Local AppData\Microsoft\Teams\Backgrounds\ folder does not exist. Trying to create it..."
        Log-Event -message $message

        try {
            New-Item -ItemType Directory -Path $TeamsBackgroundBasePath -Name "Uploads"
            $message = "Initialise-TeamsLocalUploadFolder @ "+(Get-Date)+": Successfully created Uploads folder in AppData\Microsoft\Teams\Backgrounds\."
            Log-Event -message $message

            $teamsLocalUploadFolderExists = $true
        } catch {
            $message = "Initialise-TeamsLocalUploadFolder @ "+(Get-Date)+": ERROR trying to create local Upload Folder: "+$_.Exception.Message
            Log-Event -message $message
            $teamsLocalUploadFolderExists = $false
        }
    } else {
        $teamsLocalUploadFolderExists = $true 
    }

    if ($IncludeNewTeams -eq $true) {
        $NewTeamsBackgroundBasePath = $env:LOCALAPPDATA+"\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\"
        $NewTeamsBackgroundUploadPath = $NewTeamsBackgroundBasePath+"\Uploads\"

        if (!(Test-Path $NewTeamsBackgroundUploadPath)) {
            $message = "Initialise-TeamsLocalUploadFolder @ "+(Get-Date)+": Local folder for new Teams does not exist. Indicates New Teams is not present on this system."
            Log-Event -message $message
            $NewTeamsLocalUploadFolderExists = $false
        } else {
            $NewTeamsLocalUploadFolderExists = $true
        }
    }

    return $NewTeamsLocalUploadFolderExists
}

Function Download-PictureFromURLToFolder {
    param (
        [Parameter(Mandatory=$true)] [string]$URL,
        [Parameter(Mandatory=$true)] [string]$Path
    )

    try {
        Invoke-WebRequest -Uri $URL -OutFile $Path 
        $message = "Download-PictureFromURLToFolder @ "+(Get-Date)+": Successfully Downloaded file from URL: "+$URL
        Log-Event -message $message
    } catch {
        $message = "Download-PictureFromURLToFolder @ "+(Get-Date)+": ERROR trying to download file from URL: "+$_.Exception.Message
        Log-Event -message $message
    }
}

Function Convert-FilenameToGUID {
    param(
        [Parameter(Mandatory=$true)] [string]$filenameWithoutExtension
    )

    # Compute the SHA256 hash
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($filenameWithoutExtension))

    # Convert the hash to a hexadecimal string
    $hexString = [BitConverter]::ToString($hashBytes) -replace '-', ''

    # Take the first 32 characters and format as a GUID
    $guidString = "{0}-{1}-{2}-{3}-{4}" -f $hexString.Substring(0, 8), $hexString.Substring(8, 4), $hexString.Substring(12, 4), $hexString.Substring(16, 4), $hexString.Substring(20, 12)
    $guid = [System.Guid]::Parse($guidString)

    return $guid
}

Function Get-BlobItems {
    param (
        [Parameter(Mandatory=$true)] [string]$URL
    )

    $uri = $URL.Split('?')[0]
    $sas = $URL.Split('?')[1]
    $newurl = $uri + "?restype=container&comp=list&" + $sas 

    try {
        $body = Invoke-RestMethod -Uri $newurl
        #cleanup answer and convert body to XML
        $xml = [xml]$body.Substring($body.IndexOf('<'))
        #use only the relative Path from the returned objects
        $files = $xml.ChildNodes.Blobs.Blob.Name
    } catch {
        $message = "Get-BlobItems @ "+(Get-Date)+": ERROR trying to fetch BlobItems: "+$_.Exception.Message
        Log-Event -message $message
    }
    return $files 
}

Function Set-TeamsCustomBackgrounds {
    param (
        [Parameter(Mandatory=$false)] [string]$StorageAccountName,
        [Parameter(Mandatory=$false)] [string]$ContainerName,
        [Parameter(Mandatory=$false)] [string]$SASToken,
        [Parameter(Mandatory=$false)] [boolean]$IncludeNewTeams
    )

    $BaseUrl = "https://$StorageAccountName.blob.core.windows.net/"
    $SASUrl = $BaseUrl+$ContainerName+"?"+$SASToken

    $TeamsBackgroundBasePath = $env:APPDATA+"\Microsoft\Teams\Backgrounds\"
    $TeamsBackgroundUploadPath = $TeamsBackgroundBasePath+"Uploads\"

    $NewTeamsBackgroundBasePath = $env:LOCALAPPDATA+"\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\"
    $NewTeamsBackgroundUploadPath = $NewTeamsBackgroundBasePath+"Uploads\"

    $AllFilesFromBlob = Get-BlobItems -URL $SASUrl

    foreach ($file in $AllFilesFromBlob) {
        #Process "classic" Teams
        $LocalPath = $TeamsBackgroundUploadPath+$file
        $DownloadURL = $BaseUrl+$ContainerName+"/"+$file

        if (!(Test-Path $localPath)) {
            $message = "Set-TeamsCustomBackgrounds @ "+(Get-Date)+": File "+$file+ "does not exist locally. Downloading it..."
            Log-Event -message $message
            Download-PictureFromURLToFolder -URL $DownloadURL -Path $LocalPath
        } else {
            $message = "Set-TeamsCustomBackgrounds @ "+(Get-Date)+": File "+$file+ "exists locally."
            Log-Event -message $message
        }

        #Process "New" Teams
        if ($IncludeNewTeams -eq $true) {
            $message = "Set-TeamsCustomBackgrounds @ "+(Get-Date)+": New Teams is present on this system. Checking known Upload folder for New Teams..."
            Log-Event -message $message

            #Is it a Thumbnail?
            if ($file -like "*_thumb*") {
                $FileNameWithoutExtension = ($file.Split('.'))[0].TrimEnd("_thumb")
                $GUID = (Convert-FilenameToGUID -filenameWithoutExtension $FileNameWithoutExtension).GUID
                $NewLocalPath = $NewTeamsBackgroundUploadPath+$GUID+"_thumb.jpeg"
            } else {
                $FileNameWithoutExtension = ($file.Split('.'))[0]
                $GUID = (Convert-FilenameToGUID -filenameWithoutExtension $FileNameWithoutExtension).GUID
                $NewLocalPath = $NewTeamsBackgroundUploadPath+$GUID+".jpeg"
            }

            if (!(Test-Path $NewLocalPath)) {
                $message = "Set-TeamsCustomBackgrounds @ "+(Get-Date)+": File "+$file+ "does not exist locally in: "+$NewLocalPath+". Downloading it..."
                Log-Event -message $message
                Download-PictureFromURLToFolder -URL $DownloadURL -Path $NewLocalPath
            } else {
                $message = "Set-TeamsCustomBackgrounds @ "+(Get-Date)+": File "+$file+ "exists locally in: "+$NewLocalPath
                Log-Event -message $message
            }
        }
    }
}

# Initialisation and execution
Initialise-EventLogging -LogName "Application" -Source "Intune-PoSh-SPE-WIN10-Baseline-Devices-Teams-CustomBackgrounds"

$StorageAccountName = "YOUR_STORAGE_ACCOUNT_NAME"
$ContainerName = "NAME_OF_CONTAINER"
$SASToken = "YOUR_SAS_TOKEN"

$NewTeamsIsPresent = Initialise-TeamsLocalUploadFolder -IncludeNewTeams $true

Set-TeamsCustomBackgrounds -IncludeNewTeams $NewTeamsIsPresent -StorageAccountName $StorageAccountName -ContainerName $ContainerName -SASToken $SASToken
