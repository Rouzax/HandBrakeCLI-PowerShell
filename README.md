# PowerShell Video Conversion Script

This PowerShell script allows you to convert video files in a source folder using HandBrakeCLI while providing options to test and adjust bit rates. The script provides a user-friendly menu interface and supports custom presets for the conversion process.

## Prerequisites

- PowerShell 5.1 or later
- HandBrakeCLI installed (default path: C:\Program Files\HandBrake\HandBrakeCLI.exe)
- FFmpeg installed (optional for testing; default path: C:\Program Files\FFmpeg\ffprobe.exe)

## Usage

1. Clone or download the script to your local machine.

2. Open a PowerShell terminal and navigate to the folder containing the script.

3. Run the script with the following parameters:

```powershell
.\HandBrakeCLI-PowerShell.ps1 -SourceFolder <SourceFolderPath> -OutputFolder <OutputFolderPath> [-PresetFile <PresetFilePath>] [-ConvertOnly] [-HandBrakeCLI <HandBrakeCLIPath>] [-FFprobePath <FFprobePath>] [-TestEncode] [-TestEncodeSeconds <TestDuration>]
```

### Parameters

- `SourceFolder` (mandatory): The path to the folder containing the source video files.
- `OutputFolder` (mandatory): The path where the converted files will be saved.
- `PresetFile` (optional): The path to a custom HandBrake preset JSON file. If not provided, a menu will prompt you to select a preset from the available options.
- `ConvertOnly` (optional): If specified, non-MKV files will be copied to the output folder without conversion.
- `HandBrakeCLI` (optional): The path to the HandBrakeCLI executable. If not provided, the default path will be used.
- `FFprobePath` (optional): The path to the FFprobe executable. Required only if `TestEncode` is specified. If not provided, the default path will be used.
- `TestEncode` (optional): If specified, the script will perform a test encode to check bit rates before the full conversion.
- `TestEncodeSeconds` (optional): The duration (in seconds) of the test encode. Defaults to 120 seconds.
  
## Features

- Provides a menu-based interface for easy interaction.
- Supports testing bit rates before performing a full conversion.  
- Offers flexibility in using custom presets for different conversion settings.
- Handles both MKV and non-MKV video files based on user preferences.

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
