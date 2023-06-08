param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType 'Container' })]
    [string]$SourceFolder,

    [Parameter(Mandatory = $true)]
    [string]$OutputFolder,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType 'Leaf' })]
    [string]$PresetFile,

    [Parameter()]
    [switch]$ConvertOnly,

    [string]$HandBrakeCLI
)

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
