<#
.SYNOPSIS
    Updates HandBrakeCLI to the latest version.
.DESCRIPTION
    This function checks the local version of HandBrakeCLI and updates it to the latest
    version available on GitHub if a newer version exists. It also verifies if the script
    has the necessary permissions to write to the specified installation path.
.PARAMETER HandbrakeCLIPath
    Specifies the path where HandBrakeCLI should be installed or updated. Ensure the
    script has write permissions to this location. If elevation is required, the script
    prompts the user to run with administrator privileges.
.INPUTS
    None
.OUTPUTS 
    None
.EXAMPLE
    Update-HandbrakeCLI -HandbrakeCLIPath "C:\Program Files\HandBrake\HandBrakeCLI.exe"
    # Checks and updates HandBrakeCLI to the latest version in the specified path.
.EXAMPLE
    Update-HandbrakeCLI -HandbrakeCLIPath "D:\HandBrake\HandBrakeCLI.exe"
    # Checks and updates HandBrakeCLI to the latest version in the specified path.
#>

function Update-HandbrakeCLI {
    param(
        [string]$HandbrakeCLIPath
    )

    # Display message indicating script is checking for HandBrakeCLI availability and update necessity
    Write-Host "`nChecking if HandBrakeCLI is available and update is needed" -ForegroundColor Magenta

    # Extract the folder path without the executable
    $FolderPath = Split-Path $HandbrakeCLIPath

    # Check if $MediaInfoCLIPath already exists
    if (Test-Path $HandbrakeCLIPath) {
        Write-Host "    HandBrake CLI is already installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$HandbrakeCLIPath'" -ForegroundColor Cyan  
        $installed = $true
    } else {
        # MediaInfoCLI is not installed at the specified path
        Write-Host "    HandBrake CLI is not installed at: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "'$HandbrakeCLIPath'" -ForegroundColor Cyan  
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
        # Error occurred while checking permissions
        if ($installed) {
            Write-Host "    Insufficient permissions to write to: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$FolderPath'" -ForegroundColor Cyan

            Write-Host "    Will use existing version." -ForegroundColor DarkYellow
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
    $versionFilePath = Join-Path $FolderPath "HandBrakeCLI-Version.json"

    if (Test-Path $handbrakeCLIPath) {
        # HandbrakeCLI is installed
        if (Test-Path $versionFilePath) {
            # Read the version information from the JSON file
            $versionInfo = Get-Content $versionFilePath | ConvertFrom-Json
            $currentVersion = $versionInfo.Version
            $lastLocalUpdate = $versionInfo.LastLocalUpdate
            $lastOnlineRelease = $versionInfo.LastOnlineRelease
        } else {
            # No version file found, download and get version
            $currentVersion = "0.0.0"
            $lastLocalUpdate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $lastOnlineRelease = $null

            # Save the version information to a JSON file
            $versionObject = [PSCustomObject]@{
                Version           = $currentVersion
                LastLocalUpdate   = $lastLocalUpdate
                LastOnlineRelease = $lastOnlineRelease
            }
            $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force
        }

        Write-Host "    Installed version of HandBrake CLI: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan
        Write-Host "    Last local update: " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$lastLocalUpdate" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 

        if ($null -ne $lastOnlineRelease) {
            Write-Host "    Last online release: " -ForegroundColor DarkGray -NoNewline 
            Write-Host "$lastOnlineRelease" -ForegroundColor Cyan -NoNewline 
            Write-Host "." -ForegroundColor DarkGray 
        }
    } else {
        # HandbrakeCLI is not installed
        $currentVersion = "0.0.0"
        $lastLocalUpdate = $null
        $lastOnlineRelease = $null
    }

    # Define the GitHub releases URL
    $githubApiUrl = 'https://api.github.com/repos/HandBrake/HandBrake/releases/latest'
    


    try {
        # Use Invoke-WebRequest to fetch the HTML content of the download page
        # $response = Invoke-WebRequest -Uri $MediaInfoCLIdownloadUrl -UseBasicParsing -ErrorAction Stop

        # Get the latest release information
        $latestRelease = Invoke-RestMethod -Uri $githubApiUrl
    } catch {
        if ($installed) {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$githubApiUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version." -ForegroundColor DarkYellow 
            return
        } else {
            Write-Host "    Failed to access the download page: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$githubApiUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }


    # Extract version information from the latest release
    $latestVersion = $latestRelease.tag_name
    $assets = $latestRelease.assets

    # Find the HandBrakeCLI asset with the correct name pattern
    $handbrakeCLIAsset = $assets | Where-Object { $_.name -match "HandBrakeCLI-$latestVersion-win-x86_64\.zip$" }

    if ($null -eq $handbrakeCLIAsset) {
        # Check if download links are found
        if ($installed) {
            Write-Host "    Did not find: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'HandBrakeCLI-$latestVersion-win-x86_64.zip'" -ForegroundColor Cyan

            Write-Host "    Will use existing version." -ForegroundColor DarkYellow
            return
        } else {
            Write-Host "    Did not find: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'HandBrakeCLI-$latestVersion-win-x86_64.zip'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }

    # Extract download information
    $downloadUrl = $handbrakeCLIAsset.browser_download_url

    if ($null -eq $downloadUrl) {
        # Download URL is empty
        if ($installed) {
            Write-Host "    No download links found at: " -ForegroundColor DarkYellow -NoNewline
            Write-Host "'$githubApiUrl'" -ForegroundColor Cyan

            Write-Host "    Will use existing version." -ForegroundColor DarkYellow 
            return
        } else {
            Write-Host "    No download links found at: " -ForegroundColor DarkRed -NoNewline
            Write-Host "'$githubApiUrl'" -ForegroundColor Cyan

            Write-Host "    No current version found" -ForegroundColor DarkRed
            Write-Host "    Please check your internet connection or try again later." -ForegroundColor DarkRed
            Exit
        }
    }


    # Compare versions
    if ($latestVersion -gt $currentVersion) {
        Write-Host "    Updating HandBrakeCLI from version " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion " -ForegroundColor Cyan -NoNewline 
        Write-Host "to " -ForegroundColor DarkGray -NoNewline 
        Write-Host "$latestVersion" -ForegroundColor Cyan -NoNewline 
        Write-Host "." -ForegroundColor DarkGray 

        # Define download path
        $downloadPath = Join-Path $FolderPath "HandBrakeCLI-$latestVersion-win-x86_64.zip"

        # Download the zip file
        try {
            Invoke-WebRequest -Uri $downloadUrl -OutFile $downloadPath -ErrorAction Stop
        } catch {
            if ($fileVersion) {
                Write-Host "    Failed to download HandBrakeCLI." -ForegroundColor DarkYellow

                Write-Host "    Will use existing version. " -ForegroundColor DarkYellow
                return
            } else {
                Write-Host "    Failed to download HandBrakeCLI." -ForegroundColor DarkRed

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
            Version           = $latestVersion
            LastLocalUpdate   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            LastOnlineRelease = $latestRelease.published_at
        }
        $versionObject | ConvertTo-Json | Out-File -FilePath $versionFilePath -Force

        Write-Host "    Update completed successfully." -ForegroundColor DarkGray
    } else {
        Write-Host "    HandBrakeCLI is already up to date (version :" -ForegroundColor DarkGray -NoNewline 
        Write-Host "$currentVersion" -ForegroundColor Cyan -NoNewline
        Write-Host ")." -ForegroundColor DarkGray 
    }
}

Update-HandbrakeCLI -HandbrakeCLIPath "D:\GitHub\HandBrakeCLI-PowerShell\HandBrakeCLI\HandBrakeCLI.exe"
# Update-MediaInfoCLI -MediaInfoCLIPath "C:\Program Files\HandBrake\HandBrakeCLI.exe"