param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
    [string]$PresetFile,

    [Parameter()]
    [switch]$ConvertOnly,

    [Parameter()]
    [string]$HandBrakeCLI,

    [Parameter()]
    [string]$MediaInfocliPath,

    [Parameter()]
    [switch]$TestEncode,

    [Parameter()]
    [int]$TestEncodeSeconds = 120
)


function Select-MenuOption {
    <#
    .SYNOPSIS
    This function creates a menu with options provided in the $MenuOptions parameter and prompts the user to select an option. The selected option is returned as output.
    
    .DESCRIPTION
    The Select-MenuOption function is used to create a menu with options provided as an array in the $MenuOptions parameter. The function prompts the user to select an option by displaying the options with their corresponding index numbers. The user must enter the index number of the option they wish to select. The function checks if the entered number is within the range of the available options and returns the selected option.
    
    .PARAMETER MenuOptions
    An array of options to be presented in the menu. The options must be of the same data type.
    
    .PARAMETER MenuQuestion
    A string representing the question to be asked when prompting the user for input.
    
    .EXAMPLE
    $Options = @("Option 1","Option 2","Option 3")
    $Question = "an option"
    $SelectedOption = Select-MenuOption -MenuOptions $Options -MenuQuestion $Question
    This example creates a menu with three options "Option 1", "Option 2", and "Option 3". The user is prompted to select an option by displaying the options with their index numbers. The function returns the selected option.
    
    .NOTES
    #>
    param (
        [Parameter(Mandatory = $true)]
        [Object]$MenuOptions,

        [Parameter(Mandatory = $true)]
        [string]$MenuQuestion
    )
    if ($MenuOptions.Count -eq 1) {
        Return $MenuOptions
    } else {
        Write-Host "`nSelect the correct $MenuQuestion" -ForegroundColor DarkCyan
        $menu = @{}
        $maxWidth = [math]::Ceiling([math]::Log10($MenuOptions.Count + 1))
        for ($i = 1; $i -le $MenuOptions.count; $i++) { 
            $indexDisplay = "$i.".PadRight($maxWidth + 2)
            Write-Host "$indexDisplay" -ForegroundColor Magenta -NoNewline
            Write-Host "$($MenuOptions[$i - 1])" -ForegroundColor White 
            $menu.Add($i, ($MenuOptions[$i - 1]))
        }
        do {
            try {
                $numOk = $true
                [int]$ans = Read-Host "Enter $MenuQuestion number to select"
                if ($ans -lt 1 -or $ans -gt $MenuOptions.Count) {
                    $numOK = $false
                    Write-Host 'Not a valid selection' -ForegroundColor DarkRed
                }
            } catch {
                $numOK = $false
                Write-Host 'Please enter a number' -ForegroundColor DarkRed
            }
        } # end do 
        until (($ans -ge 1 -and $ans -le $MenuOptions.Count) -and $numOK)
        Return $MenuOptions[$ans - 1]
    }
}

# Function to convert bitrate to human-readable format with two decimal places
function Convert-BitRate($bitRate) {
    <#
    .SYNOPSIS
    This function converts a given Bitrate value into a human-readable format, including bps, kbps, and Mbps.
    
    .DESCRIPTION
    The Convert-BitRate function takes a Bitrate value as input and converts it into a more readable format. It calculates and rounds the Bitrate to kilobits per second (kbps) and megabits per second (Mbps) as appropriate, and then returns the formatted result with the corresponding unit.

    .PARAMETER bitRate
    Specifies the Bitrate value that needs to be converted. It should be provided in bits per second (bps).

    .EXAMPLE
    Example 1:
    Convert-BitRate -bitRate 2500000
    This example converts a Bitrate of 2500000 bps into 2.50 Mbps.
    #>
    if ($null -eq $bitRate) {
        return ""
    }

    $kbps = [math]::Round($bitRate / 1000, 2)
    $mbps = [math]::Round($kbps / 1000, 2)

    if ($mbps -ge 1) {
        return ("{0:N2}" -f $mbps) + " Mbps"
    } elseif ($kbps -ge 1) {
        return ("{0:N2}" -f $kbps) + " kbps"
    } else {
        return "${bitRate} bps"
    }
}

# Function to extract video information using MediaInfo CLI
function Get-VideoInfo($filePath, $MediaInfocliPath) {
    <#
    .SYNOPSIS
    Retrieves detailed information about a video file using MediaInfo CLI.

    .DESCRIPTION
    This function takes a video file path and the path to the MediaInfo CLI executable as inputs.
    It uses MediaInfo to extract information about the video, such as codec, dimensions, bitrate, and encoder.

    .PARAMETER filePath
    Specifies the path to the video file for which information needs to be extracted.

    .PARAMETER MediaInfocliPath
    Specifies the path to the MediaInfo executable.

    .EXAMPLE
    Get-VideoInfo -filePath "C:\Videos\video.mp4" -MediaInfocliPath "C:\Program Files\FFmpeg\MediaInfo.exe"
    This example retrieves information about the video file "video.mp4" using MediaInfo.

    .NOTES
    This function requires MediaInfo to be installed on the system and the MediaInfocliPath parameter to point to its location.
    #>
    $MediaInfoOutput = & $MediaInfocliPath --output=JSON --Full "$filePath" | ConvertFrom-Json

    $singleVideoInfo = $null

    $generalTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'General' }
    $videoTrack = $MediaInfoOutput.media.track | Where-Object { $_.'@type' -eq 'Video' }
    
    $format = $videoTrack.Format_String
    $codec = $videoTrack.CodecID
    $videoWidth = if ($videoTrack.Width) {
        [int]$videoTrack.Width 
    } else {
        $null 
    }
    $videoHeight = if ($videoTrack.Height) {
        [int]$videoTrack.Height 
    } else {
        $null 
    }
    
    if ($videoTrack.BitRate) {
        $rawVideoBitRate = [int]$videoTrack.BitRate
        $videoBitRate = Convert-BitRate $rawVideoBitRate
    } else {
        $videoBitRate = $null
    }
    
    if ($generalTrack.OverallBitRate) {
        $rawTotalBitRate = [int]$generalTrack.OverallBitRate
        $totalBitRate = Convert-BitRate $rawTotalBitRate
    } else {
        $totalBitRate = $null
    }
    $encodedApplication = $generalTrack.Encoded_Application_String
    
    # Extracting and rounding the duration
    $rawDuration = [decimal]$videoTrack.Duration
    $videoDuration = [math]::Floor($rawDuration)

    $singleVideoInfo = [PSCustomObject]@{
        FileName        = (Get-Item -LiteralPath $filePath).BaseName
        FullPath        = $filePath
        Format          = $format
        Codec           = $codec
        VideoWidth      = $videoWidth
        VideoHeight     = $videoHeight
        VideoBitrate    = $videoBitRate
        TotalBitrate    = $totalBitRate
        VideoDuration   = $VideoDuration   
        Encoder         = $encodedApplication
        RawVideoBitrate = $rawVideoBitRate   
        RawTotalBitrate = $rawTotalBitRate  
    }
    return $singleVideoInfo
}

function Merge-VideoInfo([array]$SourceVideoInfo, [array]$TargetVideoInfo) {
    <#
    .SYNOPSIS
    Merges source and target video information based on the base name of the file and creates a combined list of video details.

    .DESCRIPTION
    This function takes two arrays of video information, source and target, and merges them based on the base name of the file. If a matching base name is found in both arrays, the function creates a combined object containing video details from both sources. If a base name is only present in the source array, its details are added as is. Finally, unmatched target base names are also included in the output.

    .PARAMETER SourceVideoInfo
    An array containing video information for source videos.

    .PARAMETER TargetVideoInfo
    An array containing video information for target videos.

    .EXAMPLE
    $sourceVideoInfo = @(
        @{ "FileName" = "video1"; "Source Format" = "H.264"; "Source Video Width" = 1920; ... },
        @{ "FileName" = "video2"; "Source Format" = "H.265"; "Source Video Width" = 1280; ... }
    )

    $targetVideoInfo = @(
        @{ "FileName" = "video1"; "Format Codec" = "H.265"; "Target Video Width" = 1920; ... },
        @{ "FileName" = "video3"; "Format Codec" = "H.264"; "Target Video Width" = 1280; ... }
    )

    $result = Merge-VideoInfo -SourceVideoInfo $sourceVideoInfo -TargetVideoInfo $targetVideoInfo

    # This example will merge video information and create a combined list containing details from both sources.
    #>

    $allVideoInfo = @()

    # Merging both tables with Source and Target video info
    foreach ($sourceVideo in $SourceVideoInfo) {
        $matchingTestVideo = $targetVideoInfo | Where-Object { $_.FileName -eq $sourceVideo.FileName }
        
        if ($matchingTestVideo) {
            $mergedObject = [PSCustomObject]@{
                FileName               = $sourceVideo.FileName
                "Source Format"        = $sourceVideo."Source Format"
                "Source Video Width"   = $sourceVideo."Source Video Width"
                "Source Video Height"  = $sourceVideo."Source Video Height"
                "Source Video Bitrate" = $sourceVideo."Source Video Bitrate"
                "Source Total Bitrate" = $sourceVideo."Source Total Bitrate"
                "Source Duration"      = $sourceVideo."Source Duration"
                "Target Format"        = $matchingTestVideo."Target Format"
                "Target Video Width"   = $matchingTestVideo."Target Video Width"
                "Target Video Height"  = $matchingTestVideo."Target Video Height"
                "Target Video Bitrate" = $matchingTestVideo."Target Video Bitrate"
                "Target Total Bitrate" = $matchingTestVideo."Target Total Bitrate"
                "Target Duration"      = $matchingTestVideo."Target Duration"
            }
            
            $allVideoInfo += $mergedObject
        } else {
            $allVideoInfo += $sourceVideo
        }
    }
    
    # Include unmatched targetVideoInfo objects
    $unmatchedtargetVideoInfo = $targetVideoInfo | Where-Object { $_.FileName -notin $SourceVideoInfo.FileName }
    $allVideoInfo += $unmatchedtargetVideoInfo

    return $allVideoInfo
}

# Set default value for $HandBrakeCLI if not provided
if (-not $HandBrakeCLI) {
    $HandBrakeCLI = "C:\Program Files\HandBrake\HandBrakeCLI.exe"
}

# Validate if HandBrakeCLI.exe exists
if (-not (Test-Path -Path $HandBrakeCLI -PathType 'Leaf')) {
    Write-Host "HandBrakeCLI.exe not found at the specified path: $HandBrakeCLI"
    Exit
}

# Set default value for $MediaInfocliPath if not provided
if (-not $MediaInfocliPath) {
    $MediaInfocliPath = "C:\Program Files\MediaInfo_CLI\MediaInfo.exe"
}
# Validate if MediaInfo.exe exists
if (-not (Test-Path -Path $MediaInfocliPath -PathType 'Leaf')) {
    Write-Host "MediaInfo.exe not found at the specified path: $MediaInfocliPath"
    Exit
}

if ($TestEncode) {
    # Do not start the full encode yet as we need to run a small test encode
    $startFullEncode = $false
} else {
    # We can directly start the entire encode
    $startFullEncode = $true
}

# Check if the OutputFolder exists and create it if not
if (-not (Test-Path -Path $OutputFolder -PathType 'Container')) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created OutputFolder: $OutputFolder"
}

# Get input if no parameters defined, list all json presets
if ($PSBoundParameters.ContainsKey('PresetFile')) {
    Write-Host "Preset File given as Parameter"
} else {
    $PresetFiles = Get-ChildItem -Path $PSScriptRoot\Presets -Filter *.json -File
    $SelectedPreset = Select-MenuOption -MenuOptions $PresetFiles -MenuQuestion "Handbrake Preset"
    $PresetFile = $SelectedPreset.FullName
}

# Read the JSON preset file
$JsonContent = Get-Content -Path $PresetFile | ConvertFrom-Json

# Get the preset name
$PresetName = $JsonContent.PresetList[0].PresetName

# Get preset extension
$VideoExtensionPreset = ($JsonContent.PresetList[0].FileFormat).Replace("av_", ".")

# Get all files in the source folder and subfolders
$allFiles = Get-ChildItem -Path $SourceFolder -File -Recurse
$allVideoFiles = $allFiles | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }

# Array to store bitrate information
$sourceVideoInfo = @()

# Get Source file information
foreach ($File in $allVideoFiles) {
    # Get the folder path and file name
    $videoInfo = Get-VideoInfo $file.FullName $MediaInfocliPath
  
    $SourceVideoInfo += [PSCustomObject]@{
        FileName               = $($videoInfo.FileName)
        "Source Format"        = $($videoInfo.Format)
        "Source Video Width"   = $($videoInfo.VideoWidth)
        "Source Video Height"  = $($videoInfo.VideoHeight)
        "Source Video Bitrate" = $($videoInfo.VideoBitrate)
        "Source Total Bitrate" = $($videoInfo.TotalBitrate)
        "Source Duration"      = $($videoInfo.VideoDuration)
    }
}

while (-not $startFullEncode) {
    # Initialize empty array 
    $targetVideoInfo = @()

    # Loop through each file and do a test encode
    foreach ($File in $allVideoFiles) {
        # Get the folder path and file name
        $SourceFilePath = $File.FullName
        $SourceFileRelativePath = $SourceFilePath.Substring($SourceFolder.Length + 1)
        # Create Output file path
        $OutputFilePath = Join-Path -Path $OutputFolder -ChildPath $SourceFileRelativePath

        # Get the current extension
        $currentExtension = [System.IO.Path]::GetExtension($OutputFilePath)
        # Replace the extension with the new one
        $OutputFilePath = $OutputFilePath -replace [regex]::Escape($currentExtension), $VideoExtensionPreset

        # Create the output folder if it doesn't exist
        $OutputFileFolder = Split-Path -Path $OutputFilePath
        if (-not (Test-Path -Path $OutputFileFolder -PathType 'Container')) {
            New-Item -ItemType Directory -Path $OutputFileFolder | Out-Null
        }

        if ($TestEncode) {

            # Get Source Video duration from 
            # Define the key you want to look up
            $lookupFileName = $File.BaseName
            # Find the object in $allVideoInfo based on the key
            $matchingObject = $SourceVideoInfo | Where-Object { $_.FileName -eq $lookupFileName }
            # Get Source Video duration
            $sourceVideoDuration = $matchingObject."Source Duration"
            
            # Calculate if a test encode with $TestEncodeSeconds can be done in the duration of the video, if yes pick a sample in the middle of the source video
            # If Source duration is shorter than $TestEncodeSeconds, pick largest possible sample
            if ($sourceVideoDuration -le $TestEncodeSeconds) {
                $startAt = 0
                $endAt = $sourceVideoDuration
            } else {
                [int]$sourceVideoMidPoint = ($sourceVideoDuration / 2)
                $startAt = $sourceVideoMidPoint - ($TestEncodeSeconds / 2)
                $endAt = $TestEncodeSeconds
            }


            # Run test encode with options
            & $HandBrakeCLI --preset-import-file "$PresetFile" --preset "$PresetName" --input "$SourceFilePath" --output "$OutputFilePath" --start-at "seconds:$startAt" --stop-at "seconds:$endAt"
        
            # Get the test video details
            $videoInfo = Get-VideoInfo $OutputFilePath $MediaInfocliPath
    
            $targetVideoInfo += [PSCustomObject]@{
                FileName               = $($videoInfo.FileName)
                "Target Format"        = $($videoInfo.Format)
                "Target Video Width"   = $($videoInfo.VideoWidth)
                "Target Video Height"  = $($videoInfo.VideoHeight)
                "Target Video Bitrate" = $($videoInfo.VideoBitrate)
                "Target Total Bitrate" = $($videoInfo.TotalBitrate)
                "Target Duration"      = $($videoInfo.VideoDuration)
            }
        }
    }

    # Initialize an empty array to store merged data
    $allVideoInfo = @()
    # Merging both tables with Source and Target video info
    $allVideoInfo = Merge-VideoInfo $SourceVideoInfo $targetVideoInfo
       
    # Show results
    Clear-Host
    Write-Host "Preset: " $PresetName
    $allVideoInfo | Format-Table -AutoSize FileName, "Source Format", "Target Format", "Source Total Bitrate", "Target Total Bitrate", "Source Video Width", "Source Video Height", "Target Video Width", "Target Video Height"
    $response = Read-Host "Is the Bitrate okay? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        $startFullEncode = $true
        
        # Clean Target Folder
        $null = Remove-Item -Path $OutputFolder -Recurse -Force
    } elseif ($response -eq 'N' -or $response -eq 'n') {
        # Prompt to select a different preset
        $SelectedPreset = Select-MenuOption -MenuOptions $PresetFiles -MenuQuestion "Handbrake Preset"
        $PresetFile = $SelectedPreset.FullName
   
        # Read the JSON preset file
        $JsonContent = Get-Content -Path $PresetFile | ConvertFrom-Json
    
        # Get the preset name
        $PresetName = $JsonContent.PresetList[0].PresetName

        # Get preset extension
        $VideoExtensionPreset = ($JsonContent.PresetList[0].FileFormat).Replace("av_", ".")

        # Clean Target Folder
        $null = Remove-Item -Path $OutputFolder -Recurse -Force

        # Don't start full encode just yet
        $startFullEncode = $false
    }
}

if ($startFullEncode) {
    # Initialize empty array 
    $targetVideoInfo = @()
    # Loop through each file
    foreach ($File in $allFiles) {
        # Get the folder path and file name
        $SourceFilePath = $File.FullName
        $SourceFileRelativePath = $SourceFilePath.Substring($SourceFolder.Length + 1)
        # Create Output file path
        $OutputFilePath = Join-Path -Path $OutputFolder -ChildPath $SourceFileRelativePath

        # Change the extension if needed on all video files
        if ($File -in $allVideoFiles) {
            # Get the current extension
            $currentExtension = [System.IO.Path]::GetExtension($OutputFilePath)
            # Replace the extension with the new one
            $OutputFilePath = $OutputFilePath -replace [regex]::Escape($currentExtension), $VideoExtensionPreset
        }

        # Create the output folder if it doesn't exist
        $OutputFileFolder = Split-Path -Path $OutputFilePath
        if (-not (Test-Path -Path $OutputFileFolder -PathType 'Container')) {
            New-Item -ItemType Directory -Path $OutputFileFolder | Out-Null
        }

        if ($File -in $allVideoFiles) {
            # Convert MKV files with HandBrakeCLI
            & $HandBrakeCLI --preset-import-file "$PresetFile" --preset "$PresetName" --input "$SourceFilePath" --output "$OutputFilePath"
            # Get the test video details
            $videoInfo = Get-VideoInfo $OutputFilePath $MediaInfocliPath
    
            $targetVideoInfo += [PSCustomObject]@{
                FileName               = $($videoInfo.FileName)
                "Target Format"        = $($videoInfo.Format)
                "Target Video Width"   = $($videoInfo.VideoWidth)
                "Target Video Height"  = $($videoInfo.VideoHeight)
                "Target Video Bitrate" = $($videoInfo.VideoBitrate)
                "Target Total Bitrate" = $($videoInfo.TotalBitrate)
                "Target Duration"      = $($videoInfo.VideoDuration)
            }
        } elseif (-not $ConvertOnly) {
            # Copy non-MKV files to the output folder
            Copy-Item -LiteralPath $SourceFilePath -Destination $OutputFilePath -Force
        }
    }
    # Initialize an empty array to store merged data
    $allVideoInfo = @()
    # Merging both tables with Source and Target video info
    $allVideoInfo = Merge-VideoInfo $SourceVideoInfo $targetVideoInfo
       
    # Show results
    Clear-Host
    Write-Host "Preset: " $PresetName
    $allVideoInfo | Format-Table -AutoSize FileName, "Source Format", "Target Format", "Source Total Bitrate", "Target Total Bitrate", "Source Video Width", "Source Video Height", "Target Video Width", "Target Video Height"
}