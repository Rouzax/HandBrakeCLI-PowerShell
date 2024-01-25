<#
.SYNOPSIS
    This script compares and encodes video files using HandBrakeCLI based on specified presets,
    providing detailed bitrate and codec information.
.DESCRIPTION
    The script processes video files in a source folder, encoding them using HandBrakeCLI with
    user-defined presets. It offers the option for a test encode to verify bitrates before
    proceeding with a full encode. The resulting video details are compared and displayed for
    evaluation.
.PARAMETER SourceFolder
    Specifies the path to the source folder containing video files for encoding.
.PARAMETER OutputFolder
    Specifies the path where the encoded video files will be saved.
.PARAMETER PresetFile
    Specifies the JSON file containing HandBrakeCLI presets. If not provided, the script
    prompts the user to select one.
.PARAMETER CopyEverything
    If present, the script will copy all files from the SourceFolder to the OutputFolder.
.PARAMETER HandBrakeCliPath
    Specifies the path to the HandBrakeCLI executable. If not provided, the default path is used.
.PARAMETER MediaInfocliPath
    Specifies the path to the MediaInfo CLI executable. If not provided, the default path is used.
.PARAMETER TestEncode
    If present, the script performs a test encode for a subset of each video to verify bitrates
    before starting the full encode.
.PARAMETER TestEncodeSeconds
    Specifies the duration, in seconds, for the test encode. Default is 120 seconds.
.EXAMPLE
    .\HandBrakeCLI-PowerShell.ps1 -SourceFolder "C:\Videos\Source" -OutputFolder "C:\Videos\Output" -PresetFile "C:\Presets\preset.json"
    This example encodes videos in the "Source" folder using the specified preset file and saves the results in the "Output" folder.
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
    [Object]$PresetFile,

    [Parameter(Mandatory = $false)]
    [switch]$CopyEverything,

    [Parameter(Mandatory = $false)]
    [string]$HandBrakeCliPath,

    [Parameter(Mandatory = $false)]
    [string]$MediaInfoCliPath,

    [Parameter(Mandatory = $false)]
    [switch]$TestEncode,

    [Parameter(Mandatory = $false)]
    [int]$TestEncodeSeconds = 120
)

<#
.SYNOPSIS
    Retrieves detailed information about video files in a specified folder, optionally recursively.
.DESCRIPTION
    This function scans a folder for video files (mp4, mkv, avi, mov, wmv) and retrieves detailed
    information using the specified MediaInfo CLI.
.PARAMETER videoFiles
    Object that holds all the video files that need to be scanned.
.PARAMETER MediaInfoCliPath
    Specifies the path to the MediaInfo CLI executable.
.PARAMETER Recursive
    Switch parameter. If present, the function scans the folder and its subfolders recursively.
.INPUTS
    Accepts file objects as input, specifically video files with extensions: mp4, mkv, avi, mov, wmv.
.OUTPUTS
    Returns an array of custom objects containing detailed information about each video file.
.EXAMPLE
    Get-VideoInfoRecursively -folderPath "C:\Videos" -MediaInfoCliPath "C:\MediaInfo\MediaInfo.exe"
.EXAMPLE
    Get-VideoInfoRecursively -folderPath "D:\Movies" -MediaInfoCliPath "D:\Tools\MediaInfo.exe" -Recursive
#>
function Get-VideoInfoRecursively {
    param (
        [Parameter(Mandatory = $true)]
        [object]$videoFiles,

        [Parameter(Mandatory = $true)]
        [string]$MediaInfoCliPath
    )

    $totalFilesToScan = ($videoFiles).Count
    $FilesScanned = 0

    $allVideoInfo = @()

    foreach ($file in $videoFiles) {
        $singleVideoInfo = $null
        $MediaInfoOutput = $null

        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $($FilesScanned + 1) of $totalFilesToScan" -Status "Reading media info: $($file.Name)" -PercentComplete $progressPercent
        
        try {
            $MediaInfoOutput = & $MediaInfoCliPath --output=JSON $file.FullName | ConvertFrom-Json
        } catch {
            Write-Host 'Exception:' $_.Exception.Message -ForegroundColor Red
            Write-Host 'Problem running MediaInfo' -ForegroundColor Red
            exit 1
        }

        # Array initialization to hold audio languages
        $audioCodecs = @()
        $audioLanguages = @()
        $audioChannels = @()

        foreach ($stream in $MediaInfoOutput.media.track) {
            # Get information from General stream
            if ($stream.'@type' -eq 'General') {
                # Get the total Bitrate
                if ($stream.OverallBitRate) {
                    [int]$rawTotalBitRate = $stream.OverallBitRate
                    $totalBitRate = Convert-BitRate -bitratePerSecond $rawTotalBitRate
                } else {
                    $totalBitRate = $null
                }

                # Get encoding Application
                [string]$encodedApplication = $stream.Encoded_Application

                # Extracting the duration
                [decimal]$rawDuration = $stream.Duration
                # Rounding video duration
                $videoDuration = [math]::Floor($rawDuration)
            }

            # Get information from Video stream
            elseif ($stream.'@type' -eq 'Video') {
                # Get Codec information from Video
                [string]$videoCodec = $stream.Format
                
                # Get Video dimensions
                [int]$videoWidth = $stream.Width 
                [int]$videoHeight = $stream.Height 
                
                # Get Video Bitrate
                if ($stream.BitRate) {
                    [int]$rawVideoBitRate = $stream.BitRate
                    $videoBitRate = Convert-BitRate -bitratePerSecond $rawVideoBitRate
                } else {
                    $videoBitRate = $null
                }
            } 

            # Get information from Audio stream
            elseif ($stream.'@type' -eq 'Audio') {
                # Keep track of all Audio Codec info
                if ($null -ne $stream.Format) {
                    $audioCodecs += $stream.Format
                } else {
                    $audioCodecs += "UND"
                }

                # Keep track of all languages
                if ($null -ne $stream.Language) {
                    $audioLanguages += $stream.Language.ToUpper()
                    
                } else {
                    $audioLanguages += "UND"
                }

                # Keep track of all Audio Channel info
                if ($null -ne $stream.Channels) {
                    $audioChannels += $stream.Channels
                } else {
                    $audioChannels += "UND"
                }
            } 

        }
                
        # Join the Codecs into a string or keep it empty if no Codecs found
        if ($audioCodecs.Count -gt 0) {
            $audioCodecs = $audioCodecs -join ' | '
        } else {
            $audioCodecs = ""
        }

        # Join the Languages into a string or keep it empty if no Languages found
        if ($audioLanguages.Count -gt 0) {
            $audioLanguages = $audioLanguages -join ' | '
        } else {
            $audioLanguages = ""
        }

        # Join the Channels into a string or keep it empty if no Channels found
        if ($audioChannels.Count -gt 0) {
            $audioChannels = $audioChannels -join ' | '
        } else {
            $audioChannels = ""
        }
        
        $singleVideoInfo = [PSCustomObject]@{
            ParentFolder    = $file.Directory.FullName
            FileName        = $file.BaseName
            FullName        = $file.FullName
            VideoCodec      = $videoCodec
            VideoWidth      = $videoWidth
            VideoHeight     = $videoHeight
            VideoBitrate    = $videoBitRate
            TotalBitrate    = $totalBitRate
            FileSize        = $(Format-Size -SizeInBytes $file.Length)
            AudioCodecs     = $audioCodecs
            AudioLanguages  = $audioLanguages
            AudioChannels   = $audioChannels
            FileSizeByte    = $file.Length
            VideoDuration   = $VideoDuration   
            Encoder         = $encodedApplication
            RawVideoBitrate = $rawVideoBitRate   
            RawTotalBitrate = $rawTotalBitRate  
        } 
                   
        if ($singleVideoInfo) {
            $allVideoInfo += $singleVideoInfo
        }

        $FilesScanned++
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Processing: $FilesScanned of $totalFilesToScan" -Status "Reading media info: $($file.Name)" -PercentComplete $progressPercent
    }
    Write-Progress -Completed -Activity "Processing: Done"
    return $allVideoInfo
}

function Start-HandBrakeCli {
    param (
        [Parameter(Mandatory = $true)]
        [object]$videoFiles,

        [Parameter(Mandatory = $true)]
        [string]$OutputFolder,

        [Parameter(Mandatory = $true)]
        $PresetFile,

        [Parameter(Mandatory = $true)]
        [string]$HandBrakeCliPath,

        [Parameter(Mandatory = $false)]
        [switch]$TestEncode,
    
        [Parameter(Mandatory = $false)]
        [int]$TestEncodeSeconds = 120
    )

    # Read the JSON preset file
    $JsonContent = Get-Content -Path $PresetFile.FullName | ConvertFrom-Json

    # Get the preset name
    $PresetName = $JsonContent.PresetList[0].PresetName

    # Get preset extension
    $VideoExtensionPreset = ($JsonContent.PresetList[0].FileFormat).Replace("av_", ".")

    $totalFilesToScan = ($videoFiles).Count
    $FilesScanned = 0

    foreach ($videoFile in $videoFiles) {
        <# $videoFile is the current item #>
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        
        # Get the folder path and file name
        $SourceFilePath = $videoFile.FullName
        $SourceFileRelativePath = $SourceFilePath.Substring($SourceFolder.Length + 1)
        # Create Output file path
        $OutputFilePath = Join-Path -Path $OutputFolder -ChildPath $SourceFileRelativePath
        
        # Create the output folder if it doesn't exist
        $OutputFileFolder = Split-Path -Path $OutputFilePath
        if (-not (Test-Path -Path $OutputFileFolder -PathType 'Container')) {
            New-Item -ItemType Directory -Path $OutputFileFolder | Out-Null
        }
        
        Write-Progress -Activity "Encoding: $($FilesScanned + 1) of $totalFilesToScan" -Status "Encoding: $($videoFile.FileName)" -PercentComplete $progressPercent
        # Change the extension if needed on all video files
        # Get the current extension
        $currentExtension = [System.IO.Path]::GetExtension($OutputFilePath)
        # Replace the extension with the new one
        $OutputFilePath = $OutputFilePath -replace [regex]::Escape($currentExtension), $VideoExtensionPreset

        # Set the base command
        $baseCommand = "& 'C:\Program Files\HandBrake\HandBrakeCLI.exe'"
        
        # Set common arguments
        $commonArguments = "--preset-import-file '$($PresetFile.Fullname)' --preset '$PresetName' --input '$SourceFilePath' --output '$OutputFilePath'"
        
        # Check if $TestEncode is $true
        if ($TestEncode) {
            # Calculate if a test encode with $TestEncodeSeconds can be done in the duration of the video, if yes pick a sample in the middle of the source video
            # If Source duration is shorter than $TestEncodeSeconds, pick largest possible sample
            $sourceVideoDuration = $videoFile.VideoDuration
            if ($sourceVideoDuration -le $TestEncodeSeconds) {
                $startAt = 0
                $endAt = $sourceVideoDuration
            } else {
                [int]$sourceVideoMidPoint = ($sourceVideoDuration / 2)
                $startAt = $sourceVideoMidPoint - ($TestEncodeSeconds / 2)
                $endAt = $TestEncodeSeconds
            }
            # Additional arguments for test encode
            $additionalArguments = "--start-at seconds:$startAt --stop-at seconds:$endAt"
        } else {
            # Default additional arguments
            $additionalArguments = ""
        }

        # Construct the full command
        $fullCommand = "$baseCommand $commonArguments $additionalArguments"
        
        
        # Execute the command
        $null = Invoke-Expression $fullCommand 2>$null
        
        $FilesScanned++
        $progressPercent = ($FilesScanned / $totalFilesToScan) * 100
        Write-Progress -Activity "Encoding: $FilesScanned of $totalFilesToScan" -Status "Encoding: $($videoFile.FileName)" -PercentComplete $progressPercent
    }
    Write-Progress -Completed -Activity "Processing: Done"
}

<#
.SYNOPSIS
    Presents a menu of options to the user and allows them to select one.
.DESCRIPTION
    This function displays a menu of options and prompts the user to select one 
    by entering the corresponding number. If there is only one option, it is 
    returned directly.
.PARAMETER MenuOptions
    Specifies an array of options to be presented in the menu.
.PARAMETER MenuQuestion
    Specifies the question or prompt to be displayed when asking the user to 
    select an option.
.OUTPUTS 
    The selected menu option.
.EXAMPLE
    $Options = @("Option 1", "Option 2", "Option 3")
    $SelectedOption = Select-MenuOption -MenuOptions $Options -MenuQuestion "an option"
    # Prompts the user to select an option and returns the selected option.
#>
function Select-MenuOption {
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

<#
.SYNOPSIS
    Converts bitrate from bits per second to a human-readable format.
.DESCRIPTION
    This function takes a bitrate value in bits per second and converts it to a
    more human-readable format. It categorizes the bitrate into Mbps, Kbps, or
    displays it in bits per second, based on the magnitude of the input value.
.PARAMETER bitratePerSecond
    Specifies the bitrate value in bits per second that needs to be converted.
    This parameter is mandatory and can accept input from the pipeline.
.INPUTS
    System.Double. Bitrate values in bits per second.
.OUTPUTS 
    System.String. The converted bitrate value along with the appropriate unit
    (Mbps, Kbps, or b/s).
.EXAMPLE
    Convert-BitRate -bitratePerSecond 1500000
    # Output: '1.50 Mb/s'
    Description: Converts the bitrate value 1500000 b/s to Mbps.
.EXAMPLE
    7500 | Convert-BitRate
    # Output: '7.50 Kb/s'
    Description: Converts the piped-in bitrate value 7500 b/s to Kbps.
.EXAMPLE
    Convert-BitRate -bitratePerSecond 500
    # Output: '500 b/s'
    Description: Displays the bitrate value 500 b/s in bits per second.
#>

function Convert-BitRate {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$bitratePerSecond
    )

    switch ($bitratePerSecond) {
        { $_ -ge 1000000 } {
            # Convert to Mb/s
            '{0:N2} Mb/s' -f ($bitratePerSecond / 1000000)
            break
        }
        { $_ -ge 1000 } {
            # Convert to Kb/s
            '{0:N2} Kb/s' -f ($bitratePerSecond / 1000)
            break
        }
        default {
            # Display in bits if less than 1 bits/s
            "$bitratePerSecond b/s"
        }
    }
}

<#
.SYNOPSIS
    Formats a size in bytes into a human-readable format.
.DESCRIPTION
    This function takes a size in bytes as input and converts it into a human-readable format, 
    displaying the size in terabytes (TB), gigabytes (GB), megabytes (MB), kilobytes (KB), 
    or bytes based on the magnitude of the input.
.PARAMETER SizeInBytes
    Specifies the size in bytes that needs to be formatted.
.INPUTS
    Accepts a double-precision floating-point number representing the size in bytes.
.OUTPUTS 
    Returns a formatted string representing the size in TB, GB, MB, KB, or bytes.
.EXAMPLE
    Format-Size -SizeInBytes 150000000000
    # Output: "139.81 GB"
    # Description: Formats 150,000,000,000 bytes into gigabytes.
.EXAMPLE
    5000000 | Format-Size
    # Output: "4.77 MB"
    # Description: Pipes 5,000,000 bytes to the function and formats the size into megabytes.
#>

function Format-Size {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipeline = $true
        )]
        [double]$SizeInBytes
    )

    switch ($SizeInBytes) {
        { $_ -ge 1PB } {
            # Convert to PB
            '{0:N2} PB' -f ($SizeInBytes / 1PB)
            break
        }
        { $_ -ge 1TB } {
            # Convert to TB
            '{0:N2} TB' -f ($SizeInBytes / 1TB)
            break
        }
        { $_ -ge 1GB } {
            # Convert to GB
            '{0:N2} GB' -f ($SizeInBytes / 1GB)
            break
        }
        { $_ -ge 1MB } {
            # Convert to MB
            '{0:N2} MB' -f ($SizeInBytes / 1MB)
            break
        }
        { $_ -ge 1KB } {
            # Convert to KB
            '{0:N2} KB' -f ($SizeInBytes / 1KB)
            break
        }
        default {
            # Display in bytes if less than 1KB
            "$SizeInBytes Bytes"
        }
    }
}

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
        @{ "FileName" = "video1"; "Source Codec" = "H.264"; "Source Video Width" = 1920; ... },
        @{ "FileName" = "video2"; "Source Codec" = "H.265"; "Source Video Width" = 1280; ... }
    )

    $targetVideoInfo = @(
        @{ "FileName" = "video1"; "Codec Codec" = "H.265"; "Target Video Width" = 1920; ... },
        @{ "FileName" = "video3"; "Codec Codec" = "H.264"; "Target Video Width" = 1280; ... }
    )

    $result = Merge-VideoInfo -SourceVideoInfo $sourceVideoInfo -TargetVideoInfo $targetVideoInfo

    # This example will merge video information and create a combined list containing details from both sources.
#>
function Merge-VideoInfo {
    param (
        [Parameter(Mandatory = $true)]
        [array]$SourceVideoInfo,

        [Parameter(Mandatory = $true)]
        [array]$TargetVideoInfo
    )

    $allVideoInfo = @()

    # Merging both tables with Source and Target video info
    foreach ($sourceVideo in $SourceVideoInfo) {
        $matchingTestVideo = $targetVideoInfo | Where-Object { $_.FileName -eq $sourceVideo.FileName }
        
        $sourceTotalRawBitrate = $sourceVideo."Source Total Raw Bitrate"
        $targetTotalRawBitrate = $matchingTestVideo."Target Total Raw Bitrate"
        
        # Calculate percentage difference
        if ($sourceTotalRawBitrate -ne 0) {
            $percentageDifference = [math]::Abs(($targetTotalRawBitrate - $sourceTotalRawBitrate) / $sourceTotalRawBitrate) * 100
        } else {
            $percentageDifference = [math]::Abs($targetTotalRawBitrate) * 100 # Handling division by zero
        }

        if ($matchingTestVideo) {
            $mergedObject = [PSCustomObject]@{
                FileName                   = $sourceVideo.FileName
                "Source Codec"             = $sourceVideo."Source Codec"
                "Source Video Width"       = $sourceVideo."Source Video Width"
                "Source Video Height"      = $sourceVideo."Source Video Height"
                "Source Video Bitrate"     = $sourceVideo."Source Video Bitrate"
                "Source Total Bitrate"     = $sourceVideo."Source Total Bitrate"
                "Source Total Raw Bitrate" = $sourceVideo."Source Total Raw Bitrate"
                'Reduction in Bitrate %'   = $percentageDifference
                "Source Duration"          = $sourceVideo."Source Duration"
                "Target Codec"             = $matchingTestVideo."Target Codec"
                "Target Video Width"       = $matchingTestVideo."Target Video Width"
                "Target Video Height"      = $matchingTestVideo."Target Video Height"
                "Target Video Bitrate"     = $matchingTestVideo."Target Video Bitrate"
                "Target Total Bitrate"     = $matchingTestVideo."Target Total Bitrate"
                "Target Total Raw Bitrate" = $matchingTestVideo."Target Total Raw Bitrate"
                "Target Duration"          = $matchingTestVideo."Target Duration"
            }
            
            $allVideoInfo += $mergedObject
        } else {
            $allVideoInfo += $sourceVideo
        }
    }
    
    # Include unmatched targetVideoInfo objects
    $unmatchedTargetVideoInfo = $targetVideoInfo | Where-Object { $_.FileName -notin $SourceVideoInfo.FileName }
    $allVideoInfo += $unmatchedTargetVideoInfo

    return $allVideoInfo
}

# Handle no MediaInfocliPath Path given as parameter
if (-not $PSBoundParameters.ContainsKey('MediaInfocliPath')) {
    $MediaInfocliPath = (Get-Command MediaInfo.exe -ErrorAction SilentlyContinue).Path 
    if (-not $MediaInfocliPath) {
        $MediaInfocliPath = "C:\Program Files\MediaInfo_CLI\MediaInfo.exe"
    }
}
# Check if MediaInfo executable exists
if (-not (Test-Path $MediaInfocliPath)) {
    Write-Host "Error: MediaInfo CLI executable not found at the specified path: $MediaInfocliPath"
    Write-Host "Please provide the correct path to the MediaInfo CLI executable using the -MediaInfocliPath parameter."
    Exit
}

# Handle no HandBrakeCLI Path given as parameter
if (-not $PSBoundParameters.ContainsKey('HandBrakeCliPath')) {
    $HandBrakeCliPath = (Get-Command HandBrakeCLI.exe -ErrorAction SilentlyContinue).Path 
    if (-not $HandBrakeCliPath) {
        $HandBrakeCliPath = "C:\Program Files\HandBrake\HandBrakeCLI.exe"
    }
}
# Check if MediaInfo executable exists
if (-not (Test-Path $HandBrakeCliPath)) {
    Write-Host "Error: HandBrake CLI executable not found at the specified path: $HandBrakeCliPath"
    Write-Host "Please provide the correct path to the MediaInfo CLI executable using the HandBrakeCliPath parameter."
    Exit
}

# Check if the OutputFolder exists and create it if not
if (-not (Test-Path -Path $OutputFolder -PathType 'Container')) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created OutputFolder: $OutputFolder"
}

# Get input if no parameters defined, list all json presets
if ($PSBoundParameters.ContainsKey('PresetFile')) {
    Write-Host "Preset File given as Parameter"
    $PresetFile = Get-Item -Path $PresetFile
} else {
    $PresetFiles = Get-ChildItem -Path $PSScriptRoot\Presets -Filter *.json -File
    $SelectedPreset = Select-MenuOption -MenuOptions $PresetFiles.BaseName -MenuQuestion "Handbrake Preset"
    $PresetFile = $PresetFiles | Where-Object { $_.BaseName -match $SelectedPreset }
}

if ($TestEncode) {
    # Do not start the full encode yet as we need to run a small test encode
    $startFullEncode = $false
} else {
    # We can directly start the entire encode
    $startFullEncode = $true
}

# Initialize arrays
$allFiles = @()
$sourceVideoObj = @()
$sourceVideoInfo = @()
$sourceVideoFiles = @()


#* Start of script
##Clear-Host

$FilesParams = @{
    Recurse = $true
    Path    = $SourceFolder
    File    = $true
}
$allFiles = Get-ChildItem @FilesParams
$sourceVideoFiles = $allFiles | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }
$sourceNonVideoFiles = $allFiles | Where-Object { $_.Extension -notmatch '\.(mp4|mkv|avi|mov|wmv)$' }

if ($sourceVideoFiles.count -eq 0) {
    Write-Host "No video files found in location: $SourceFolder"
    Exit
}

# Get source video information
$sourceVideoInfo = Get-VideoInfoRecursively -videoFiles $sourceVideoFiles -MediaInfoCliPath $MediaInfocliPath
foreach ($videoFile in $sourceVideoInfo) {
    # Construct an object to hold the values
    $SourceVideoObj += [PSCustomObject]@{
        FileName                   = $($videoFile.FileName)
        "Source FullName"          = $($videoFile.FullName)
        "Source Codec"             = $($videoFile.VideoCodec)
        "Source Video Width"       = $($videoFile.VideoWidth)
        "Source Video Height"      = $($videoFile.VideoHeight)
        "Source Video Bitrate"     = $($videoFile.VideoBitrate)
        "Source Total Bitrate"     = $($videoFile.TotalBitrate)
        "Source Total Raw Bitrate" = $($videoFile.RawTotalBitrate)
        "Source Duration"          = $($videoFile.VideoDuration)
    }
}

while (-not $startFullEncode) {
    # Initialize arrays
    $targetVideoObj = @()
    $targetVideoInfo = @()
    $targetVideoFiles = @()
    
    # Start the test encodes
    Start-HandBrakeCli -videoFiles $sourceVideoInfo -OutputFolder $OutputFolder -PresetFile $PresetFile -HandBrakeCliPath $HandBrakeCliPath -TestEncode -TestEncodeSeconds $TestEncodeSeconds

    $FilesParams = @{
        Recurse = $true
        Path    = $OutputFolder
        File    = $true
    }
    $targetVideoFiles = Get-ChildItem @FilesParams | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }

    # Get Target video information
    $targetVideoInfo = Get-VideoInfoRecursively -videoFiles $targetVideoFiles -MediaInfoCliPath $MediaInfocliPath
    foreach ($videoFile in $targetVideoInfo) {
        # Construct an object to hold the values
        $targetVideoObj += [PSCustomObject]@{
            FileName                   = $($videoFile.FileName)
            "Target FullName"          = $($videoFile.FullName)
            "Target Codec"             = $($videoFile.VideoCodec)
            "Target Video Width"       = $($videoFile.VideoWidth)
            "Target Video Height"      = $($videoFile.VideoHeight)
            "Target Video Bitrate"     = $($videoFile.VideoBitrate)
            "Target Total Bitrate"     = $($videoFile.TotalBitrate)
            "Target Total Raw Bitrate" = $($videoFile.RawTotalBitrate)
            "Target Duration"          = $($videoFile.VideoDuration)
        }
    }

    # Initialize an empty array to store merged data
    $combinedVideoInfo = @()
    # Merging both tables with Source and Target video info
    $combinedVideoInfo = Merge-VideoInfo $SourceVideoObj $targetVideoObj
    # Show results
   Clear-Host
    Write-Host "Preset: " $PresetName
    $combinedVideoInfo | Select-Object -Property FileName, "Source Codec", "Target Codec", "Source Total Bitrate", "Target Total Bitrate", 'Reduction in Bitrate %', "Source Video Width", "Source Video Height", "Target Video Width", "Target Video Height" | Out-GridView -Title "Compare Source and Test Target properties"
    $response = Read-Host "Is the Bitrate okay? (Y/N)"

    if ($response -eq 'Y' -or $response -eq 'y') {
       Clear-Host
        # We are happy with the test encode, full encode can start
        $startFullEncode = $true
        
        # Clean Target Folder
        $null = Remove-Item -Path $OutputFolder -Recurse -Force
    } elseif ($response -eq 'N' -or $response -eq 'n') {
       Clear-Host
        # Prompt to select a different preset
        $SelectedPreset = Select-MenuOption -MenuOptions $PresetFiles.BaseName -MenuQuestion "Handbrake Preset"
        $PresetFile = $PresetFiles | Where-Object { $_.BaseName -match $SelectedPreset }

        # Clean Target Folder
        $null = Remove-Item -Path $OutputFolder -Recurse -Force
        
        # Don't start full encode just yet
        $startFullEncode = $false
    }
}

if ($startFullEncode) {
    # Initialize arrays
    $targetVideoObj = @()
    $targetVideoInfo = @()
    $targetVideoFiles = @()

    # Start the Full encodes
    Start-HandBrakeCli -videoFiles $sourceVideoInfo -OutputFolder $OutputFolder -PresetFile $PresetFile -HandBrakeCliPath $HandBrakeCliPath

    $FilesParams = @{
        Recurse = $true
        Path    = $OutputFolder
        File    = $true
    }
    $targetVideoFiles = Get-ChildItem @FilesParams | Where-Object { $_.Extension -match '\.(mp4|mkv|avi|mov|wmv)$' }

    # Get Target video information
    $targetVideoInfo = Get-VideoInfoRecursively -videoFiles $targetVideoFiles -MediaInfoCliPath $MediaInfocliPath
    foreach ($videoFile in $targetVideoInfo) {
        # Construct an object to hold the values
        $targetVideoObj += [PSCustomObject]@{
            FileName                   = $($videoFile.FileName)
            "Target FullName"          = $($videoFile.FullName)
            "Target Codec"             = $($videoFile.VideoCodec)
            "Target Video Width"       = $($videoFile.VideoWidth)
            "Target Video Height"      = $($videoFile.VideoHeight)
            "Target Video Bitrate"     = $($videoFile.VideoBitrate)
            "Target Total Bitrate"     = $($videoFile.TotalBitrate)
            "Target Total Raw Bitrate" = $($videoFile.RawTotalBitrate)
            "Target Duration"          = $($videoFile.VideoDuration)
        }
    }

    if ($CopyEverything) {
        $totalFilesToCopy = ($sourceNonVideoFiles).Count
        $FilesCopied = 0
        foreach ($file in $sourceNonVideoFiles) {
            <# $file is the current item #>
            $progressPercent = ($FilesCopied / $totalFilesToCopy) * 100
            
            # Get the folder path and file name
            $SourceFilePath = $file.FullName
            $SourceFileRelativePath = $SourceFilePath.Substring($SourceFolder.Length + 1)
            # Create Output file path
            $OutputFilePath = Join-Path -Path $OutputFolder -ChildPath $SourceFileRelativePath
            Write-Progress -Activity "Copy: $($FilesCopied + 1) of $totalFilesToCopy" -Status "File: $($file.FileName)" -PercentComplete $progressPercent
            
            # Copy non-MKV files to the output folder
            Copy-Item -LiteralPath $SourceFilePath -Destination $OutputFilePath -Force
            
            $FilesCopied++
            $progressPercent = ($FilesCopied / $totalFilesToCopy) * 100
            Write-Progress -Activity "Copy: $FilesCopied of $totalFilesToCopy" -Status "File: $($file.FileName)" -PercentComplete $progressPercent
        }
        Write-Progress -Completed -Activity "Processing: Done"
     
    }

    # Initialize an empty array to store merged data
    $combinedVideoInfo = @()
    # Merging both tables with Source and Target video info
    $combinedVideoInfo = Merge-VideoInfo $SourceVideoObj $targetVideoObj
    # Show results
   Clear-Host
    Write-Host "Preset: " $PresetName
    $combinedVideoInfo | Select-Object -Property FileName, "Source Codec", "Target Codec", "Source Total Bitrate", "Target Total Bitrate", 'Reduction in Bitrate %', "Source Video Width", "Source Video Height", "Target Video Width", "Target Video Height" | Out-GridView -Title "Compare Source and Test Target properties"
}

