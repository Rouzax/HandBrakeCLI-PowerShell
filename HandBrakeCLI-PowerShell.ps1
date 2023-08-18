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
    [string]$FFprobePath,

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
        for ($i = 1; $i -le $MenuOptions.count; $i++) { 
            Write-Host "$i. " -ForegroundColor Magenta -NoNewline
            Write-Host "$($MenuOptions[$i - 1])" -ForegroundColor White 
            $menu.Add($i, ($MenuOptions[$i - 1]))
            $maxi++
        }
        do {
            try {
                $numOk = $true
                [int]$ans = Read-Host "Enter $MenuQuestion number to select"
                if ($ans -lt 1 -or $ans -gt $maxi) {
                    $numOK = $false
                    Write-Host 'Not a valid selection' -ForegroundColor DarkRed
                }
            } catch {
                $numOK = $false
                Write-Host 'Please enter a number' -ForegroundColor DarkRed
            }
        } # end do 
        until (($ans -ge 1 -and $ans -le $maxi) -and $numOK)
        Return $MenuOptions[$ans - 1]
    }
}

# Function to convert bitrate to human-readable format with two decimal places
function Convert-BitRate($bitRate) {
    <#
    .SYNOPSIS
    This function converts a given bit rate value into a human-readable format, including bps, kbps, and Mbps.
    
    .DESCRIPTION
    The Convert-BitRate function takes a bit rate value as input and converts it into a more readable format. It calculates and rounds the bit rate to kilobits per second (kbps) and megabits per second (Mbps) as appropriate, and then returns the formatted result with the corresponding unit.

    .PARAMETER bitRate
    Specifies the bit rate value that needs to be converted. It should be provided in bits per second (bps).

    .EXAMPLE
    Example 1:
    Convert-BitRate -bitRate 2500000
    This example converts a bit rate of 2500000 bps into 2.50 Mbps.
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

# Set default value for $HandBrakeCLI if not provided
if (-not $HandBrakeCLI) {
    $HandBrakeCLI = "C:\Program Files\HandBrake\HandBrakeCLI.exe"
}

# Validate if HandBrakeCLI.exe exists
if (-not (Test-Path -Path $HandBrakeCLI -PathType 'Leaf')) {
    Write-Host "HandBrakeCLI.exe not found at the specified path: $HandBrakeCLI"
    Exit
}

if ($TestEncode) {
    if (-not $FFprobePath) {
        $FFprobePath = "C:\Program Files\FFmpeg\ffprobe.exe"
    }
    # Validate if FFprobe.exe exists
    if (-not (Test-Path -Path $FFprobePath -PathType 'Leaf')) {
        Write-Host "FFprobe.exe not found at the specified path: $FFprobePath"
        Exit
    }
} else {
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

# Get all files in the source folder and subfolders
$Files = Get-ChildItem -Path $SourceFolder -File -Recurse


while (-not $startFullEncode) {
    # Array to store bitrate information
    $bitRateInfo = @()
    # Loop through each file
    foreach ($File in $Files) {
        # Get the folder path and file name
        $SourceFilePath = $File.FullName
        $SourceFileRelativePath = $SourceFilePath.Substring($SourceFolder.Length + 1)
        $OutputFilePath = Join-Path -Path $OutputFolder -ChildPath $SourceFileRelativePath

        # Create the output folder if it doesn't exist
        $OutputFileFolder = Split-Path -Path $OutputFilePath
        if (-not (Test-Path -Path $OutputFileFolder -PathType 'Container')) {
            New-Item -ItemType Directory -Path $OutputFileFolder | Out-Null
        }

        if ($TestEncode) {
            $startFullEncode = $false

            # Use FFprobe to get Source video information
            $ffprobeOutput = & $FFprobePath -v error -print_format json -show_format -show_streams "$SourceFilePath" | ConvertFrom-Json
            $sourceVideoDuration = [int]$ffprobeOutput.format.duration
            if ($sourceVideoDuration -le $TestEncodeSeconds) {
                $startAt = 0
                $endAt = $sourceVideoDuration
            } else {
                [int]$sourceVideoMidPoint = ($sourceVideoDuration / 2)
                $startAt = $sourceVideoMidPoint - ($TestEncodeSeconds / 2)
                $endAt = $TestEncodeSeconds
            }
            $sourceBitRate = $ffprobeOutput.format.bit_rate
            $sourceBitRateFormatted = Convert-BitRate $sourceBitRate

            # Run test encode with options
            & $HandBrakeCLI --preset-import-file "$PresetFile" --preset "$PresetName" --input "$SourceFilePath" --output "$OutputFilePath" --start-at "seconds:$startAt" --stop-at "seconds:$endAt"
        
            # Use FFprobe to get bit rate of test encode
            $ffprobeTestOutput = & $FFprobePath -v error -print_format json -show_format "$OutputFilePath" | ConvertFrom-Json
            $testBitRate = $ffprobeTestOutput.format.bit_rate
            $testBitRateFormatted = Convert-BitRate $testBitRate

            $bitRateInfo += [PSCustomObject]@{
                FileName          = $File.Name
                "Source Bit Rate" = $sourceBitRateFormatted
                "Test Bit Rate"   = $testBitRateFormatted
            }
        }
    }

    Clear-Host
    
    Write-Host "Preset: " $PresetName
    $bitRateInfo | Format-Table -AutoSize FileName, "Source Bit Rate", "Test Bit Rate"
    $response = Read-Host "Is the Bit Rate okay? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        $startFullEncode = $true
    } elseif ($response -eq 'N' -or $response -eq 'n') {
        # Prompt to select a different preset
        $SelectedPreset = Select-MenuOption -MenuOptions $PresetFiles -MenuQuestion "Handbrake Preset"
        $PresetFile = $SelectedPreset.FullName
   
        # Read the JSON preset file
        $JsonContent = Get-Content -Path $PresetFile | ConvertFrom-Json
    
        # Get the preset name
        $PresetName = $JsonContent.PresetList[0].PresetName

        # Don't start full encode just yet
        $startFullEncode = $false
    }
}
if ($startFullEncode) {
    if (($File.Extension -eq '.mkv') -or ($File.Extension -eq '.mp4')) {
        # Convert MKV files with HandBrakeCLI
        & $HandBrakeCLI --preset-import-file "$PresetFile" --preset "$PresetName" --input "$SourceFilePath" --output "$OutputFilePath"
    } elseif (-not $ConvertOnly) {
        # Copy non-MKV files to the output folder
        Copy-Item -LiteralPath $SourceFilePath -Destination $OutputFilePath -Force
    }
}