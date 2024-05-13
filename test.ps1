<#
.SYNOPSIS
    Updates MediaInfoCLI to the latest version.
.DESCRIPTION
    This function checks the local version of MediaInfoCLI and updates it to the latest
    version available on GitHub if a newer version exists. It also verifies if the script
    has the necessary permissions to write to the specified installation path.
.PARAMETER MediaInfoCLIPath
    Specifies the path where MediaInfoCLI should be installed or updated. Ensure the
    script has write permissions to this location. If elevation is required, the script
    prompts the user to run with administrator privileges.
.INPUTS
    None
.OUTPUTS 
    None
.EXAMPLE
    Update-MediaInfoCLI -MediaInfoCLIPath "C:\Program Files\HandBrake\MediaInfoCLI.exe"
    # Checks and updates MediaInfoCLI to the latest version in the specified path.
.EXAMPLE
    Update-MediaInfoCLI -MediaInfoCLIPath "D:\HandBrake\MediaInfoCLI.exe"
    # Checks and updates MediaInfoCLI to the latest version in the specified path.
#>
function Update-MediaInfoCLI {
    param(
        [string]$MediaInfoCLIPath
    )
    Write-Host "`nChecking if MediaInfo CLI is available and update is needed" -ForegroundColor Magenta
  
    # Extract the folder path without the executable
    $FolderPath = Split-Path $MediaInfoCLIPath
    
    # Check if $MediaInfoCLIPath already exists
    if (Test-Path $MediaInfoCLIPath) {
        Write-Host "    MediaInfo CLI is already installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$MediaInfoCLIPath'" -ForegroundColor Cyan  
        $fileVersion = (Get-Command $MediaInfoCLIPath).FileVersionInfo.ProductVersion
    } else {
        # MediaInfoCLI is not installed at the specified path
        Write-Host "    MediaInfo CLI is not installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$MediaInfoCLIPath'" -ForegroundColor Cyan  
        Write-Host "    Will try to get the latest version online" -ForegroundColor DarkGray 
    }
    
    # Check if writing to $FolderPath requires elevated permissions
    try {
        # Check if the folder exists, if not create it
        if (-not (Test-Path $FolderPath)) {
            New-Item -ItemType Directory -Path $FolderPath -ErrorAction Stop | Out-Null
        }

        # Check if writing to $FolderPath requires elevated permissions
        $testPath = Join-Path $FolderPath "Test-Permission.txt"
        $null | Out-File -FilePath $testPath -Force -ErrorAction Stop
        Remove-Item -Path $testPath -Force -ErrorAction Stop
    } catch {
        if ($fileVersion) {
            Write-Host "    Insufficient permissions to write to: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$FolderPath'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Insufficient permissions to write to: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$FolderPath'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please run the script as an administrator." -ForegroundColor DarkRed
            Exit
        }
    }

    # Define version file path
    $versionFilePath = Join-Path $FolderPath "MediaInfoCLI-Version.json"

    if (Test-Path $MediaInfoCLIPath) {
        # MediaInfoCLI is installed
        if (Test-Path $versionFilePath) {
            # Read the version information from the JSON file
            $versionInfo = Get-Content $versionFilePath | ConvertFrom-Json
            $currentVersion = $versionInfo.Version
            $lastLocalUpdate = $versionInfo.LastLocalUpdate
        } else {
            # No version file found, write current version and date to MediaInfoCLI-Version.json
            $lastLocalUpdate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $currentVersion = $fileVersion
            # Save the version information to a JSON file
            $versionObject = [PSCustomObject]@{
                Version         = $currentVersion
                LastLocalUpdate = $lastLocalUpdate
            }
            $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force
        }
        
        # Display installed version and last update date
        Write-Host "    Installed version of MediaInfo CLI: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan
        Write-Host "    Last local update: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$lastLocalUpdate" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 
    } else {
        # MediaInfoCLI is not installed
        $currentVersion = "0.0"
        $lastLocalUpdate = $null
    }

    # Define the download URL for MediaInfo_CLI
    $MediaInfoCLIdownloadUrl = "https://mediaarea.net/en/MediaInfo/Download/Windows"

    try {
        # Use Invoke-WebRequest to fetch the HTML content of the download page
        $response = Invoke-WebRequest -Uri $MediaInfoCLIdownloadUrl -UseBasicParsing -ErrorAction Stop
    } catch {
        if ($fileVersion) {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Parse the HTML response to extract the download links
    $downloadLinks = $response.Links | Where-Object { $_.href -match "/MediaInfo_CLI_[\d.]+_Windows_x64.zip" }

    # Check if download links are found
    if ($downloadLinks.Count -eq 0) {
        if ($fileVersion) {
            Write-Host "    No download links found at: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Did not find: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    No download links found at: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Did not find: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Initialize variables to store version and latest link
    $latestVersion = $null
    $latestLink = $null

    # Iterate through each download link to find the latest version
    foreach ($link in $downloadLinks) {
        if ($link.href -match '/MediaInfo_CLI_([\d.]+)_Windows_x64.zip') {
            $version = $Matches[1]
            if (-not $latestVersion -or [version]::Parse($version) -gt [version]::Parse($latestVersion)) {
                $latestVersion = $version
                $latestLink = $link
            }
        }
    }
    
    # Check if latest link is found
    if (-not $latestLink) {
        if ($fileVersion) {
            Write-Host "    Did not find: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    Did not find: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'MediaInfo_CLI_(x.x)_Windows_x64.zip'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Construct the full download URL
    $downloadUrl = "https:" + $latestLink.href
    
    if ($null -eq $downloadUrl) {
         # Download URL is empty
        if ($fileVersion) {
            Write-Host "    No download links found at: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
            Write-Host $fileVersion -ForegroundColor Cyan
            return
        } else {
            Write-Host "    No download links found at: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$MediaInfoCLIdownloadUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Compare versions
    if ($latestVersion -gt $currentVersion) {
        # Newer version available, proceed with update

        # Display updating message
        Write-Host "    Updating MediaInfoCLI from version " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion " -ForegroundColor Cyan -NoNewline 
        Write-Host "to " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$latestVersion" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 

        # Define download path
        $downloadPath = Join-Path $FolderPath "MediaInfo_CLI_$($latestVersion)_Windows_x64.zip"

        # Download the zip file
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -ErrorAction Stop
        } catch {
            if ($fileVersion) {
                Write-Host "    Failed to download MediaInfoCLI." -ForegroundColor DarkYellow
    
                Write-Host "    Will use existing version: " -ForegroundColor DarkYellow -NoNewline
                Write-Host $fileVersion -ForegroundColor Cyan
                return
            } else {
                Write-Host "    Failed to download MediaInfoCLI." -ForegroundColor DarkRed
    
                Write-Host "    No current version found" -ForegroundColor DarkRed
                Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
                Exit
            }
        }

        # Extract the contents
        Write-Host "    Extracting files to " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$FolderPath" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 
        try {
            Expand-Archive -Path $downloadPath -DestinationPath $FolderPath -Force -ErrorAction Stop
        } catch {
            Write-Host "    Failed to extract files. Aborting update." -ForegroundColor DarkRed
            Write-Host "    Error: $_" -ForegroundColor DarkRed
            exit
        }

        # Clean up the downloaded zip file
        Remove-Item -Path $downloadPath -Force

        # Update the version information in the JSON file
        $versionObject = [PSCustomObject]@{
            Version         = $latestVersion
            LastLocalUpdate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force

        Write-Host "    Update completed successfully." -ForegroundColor DarkGray
    } else {
        Write-Host "    MediaInfoCLI is already up to date (version :" -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan -NoNewline
        Write-Host ")." -ForegroundColor DarkGray 
    }
}

Update-MediaInfoCLI -MediaInfoCLIPath "D:\GitHub\HandBrakeCLI-PowerShell\MediaInfoCLI\MediaInfo.exe"
# Update-MediaInfoCLI -MediaInfoCLIPath "C:\Program Files\MediaInfo_CLI\MediaInfo.exe"