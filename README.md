
# PowerShell Script: File Conversion and Copy Utility

This PowerShell script is designed to convert MKV files using HandBrakeCLI and copy non-MKV files from a source folder to an output folder. It provides flexibility in specifying the source folder, output folder, preset file, and HandBrakeCLI executable.

## Prerequisites

- PowerShell 5.1 or higher
- HandBrakeCLI executable (default path: "C:\Program Files\HandBrake\HandBrakeCLI.exe")

## Usage

1. Ensure you have the required prerequisites mentioned above.
2. Download or clone the PowerShell script from the repository.
3. Open a PowerShell command prompt.
4. Navigate to the directory where the script is located.
5. Run the script using the following command:

   ```powershell
   .\ConvertAndCopyFiles.ps1 -SourceFolder <SourceFolderPath> -OutputFolder <OutputFolderPath> -PresetFile <PresetFilePath> [-ConvertOnly] [-HandBrakeCLI <HandBrakeCLIPath>]
   ```

   Replace the placeholders `<SourceFolderPath>`, `<OutputFolderPath>`, `<PresetFilePath>`, and `<HandBrakeCLIPath>` with the actual paths.

6. If the `-ConvertOnly` switch is provided, the script will only convert MKV files using HandBrakeCLI. Otherwise, it will copy non-MKV files to the output folder as well.

7. The script will create the output folder if it doesn't exist.

## Parameters

- `SourceFolder`: The path to the source folder containing the files to be processed.
- `OutputFolder`: The path to the output folder where the converted files will be stored.
- `PresetFile`: The path to the JSON preset file for HandBrakeCLI.
- `ConvertOnly` (optional): A switch parameter indicating whether to convert only or also copy non-MKV files.
- `HandBrakeCLI` (optional): The path to the HandBrakeCLI executable. If not specified, the default path will be used.

## Example

Convert and copy files from "D:\Movies\Source" to "D:\Movies\Output" using a preset file "D:\Temp\handbrake\preset.json":

```powershell
.\ConvertAndCopyFiles.ps1 -SourceFolder "D:\Movies\Source" -OutputFolder "D:\Movies\Output" -PresetFile "D:\Temp\handbrake\preset.json"
```

Convert MKV files only from "D:\Movies\Source" to "D:\Movies\Output" using a preset file "D:\Temp\handbrake\preset.json":

```powershell
.\ConvertAndCopyFiles.ps1 -SourceFolder "D:\Movies\Source" -OutputFolder "D:\Movies\Output" -PresetFile "D:\Temp\handbrake\preset.json" -ConvertOnly
```