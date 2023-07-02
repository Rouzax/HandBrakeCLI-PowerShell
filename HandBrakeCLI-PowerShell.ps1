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

    [string]$HandBrakeCLI
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

# Set default value for $HandBrakeCLI if not provided
if (-not $HandBrakeCLI) {
    $HandBrakeCLI = "C:\Program Files\HandBrake\HandBrakeCLI.exe"
}

# Check if the OutputFolder exists and create it if not
if (-not (Test-Path -Path $OutputFolder -PathType 'Container')) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    Write-Host "Created OutputFolder: $OutputFolder"
}

# Validate if HandBrakeCLI.exe exists
if (-not (Test-Path -Path $HandBrakeCLI -PathType 'Leaf')) {
    Write-Host "HandBrakeCLI.exe not found at the specified path: $HandBrakeCLI"
    Exit
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

    if ($File.Extension -eq '.mkv') {
        # Convert MKV files with HandBrakeCLI
        & $HandBrakeCLI --preset-import-file "$PresetFile" --preset "$PresetName" --input "$SourceFilePath" --output "$OutputFilePath"
    } elseif (-not $ConvertOnly) {
        # Copy non-MKV files to the output folder
        Copy-Item -LiteralPath $SourceFilePath -Destination $OutputFilePath -Force
    }
}
