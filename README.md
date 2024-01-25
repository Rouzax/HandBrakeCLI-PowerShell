# PowerShell Video Conversion Script

This script is designed to automate the process of converting video files using HandBrakeCLI and MediaInfo CLI. It supports various options for source and output folders, preset selection, test encoding, and more. The script is primarily intended for Windows environments.

## Features

- Convert video files using HandBrakeCLI
- Extract video information using MediaInfo CLI
- Test encoding to check video quality and bitrate
- Select presets for conversion
- Merge source and target video information for comparison


## Requirements

- [HandBrakeCLI](https://handbrake.fr/downloads2.php)
- [MediaInfo CLI](https://mediaarea.net/en/MediaInfo/Download)

## Usage

1. Place your video files in the source folder.
2. Run the script in PowerShell, providing the required parameters.
3. The script will perform a test encode to ensure desired bitrates if `-TestEncode` is provided.
4. If the test results are satisfactory, the script will perform the full conversion.

```powershell
.\HandBrakeCLI-PowerShell.ps1 -SourceFolder <SourceFolderPath> -OutputFolder <OutputFolderPath> [-PresetFile <PresetFilePath>] [-CopyEverything] [-HandBrakeCliPath <HandBrakeCLIPath>] [-TestEncode] [-TestEncodeSeconds <TestDuration>]
```

### Parameters

- `SourceFolder` (mandatory): The path to the folder containing the source video files.
- `OutputFolder` (mandatory): The path where the converted files will be saved.
- `PresetFile` (optional): The path to a custom HandBrake preset JSON file. If not provided, a menu will prompt you to select a preset from the available options.
- `CopyEverything` (optional): If present, the script will copy all files from the SourceFolder to the OutputFolder.
- `HandBrakeCliPath` (optional): The path to the HandBrakeCLI executable. If not provided, the default path will be used.
- `MediaInfocliPath` (Optional): Path to MediaInfo CLI executable.
- `TestEncode` (optional): If specified, the script will perform a test encode to check bit rates before the full conversion.
- `TestEncodeSeconds` (optional): The duration (in seconds) of the test encode. Defaults to 120 seconds.
  
## Example

Convert and copy files from "D:\Movies\Source" to "D:\Movies\Output":

```powershell
.\HandBrakeCLI-PowerShell.ps1 -SourceFolder "D:\Movies\Source" -OutputFolder "D:\Movies\Output"
```

Convert MKV files only from "D:\Movies\Source" to "D:\Movies\Output":

```powershell
.\HandBrakeCLI-PowerShell.ps1 -SourceFolder "D:\Movies\Source" -OutputFolder "D:\Movies\Output" -ConvertOnly
```

Convert video files from the ""D:\Movies\Source" folder to the "D:\Movies\Output" folder. It will perform a 60-second test encode for bit rate evaluation.

```powershell
.\HandBrakeCLI-PowerShell.ps1 -SourceFolder "D:\Movies\Source" -OutputFolder "D:\Movies\Output" -TestEncode -TestEncodeSeconds 60
```

## Presets

The script supports using custom HandBrake presets in JSON format. Place your preset files in the "Presets" folder located in the same directory as the script. When prompted to select a preset file, choose from the available options in the "Presets" folder.
