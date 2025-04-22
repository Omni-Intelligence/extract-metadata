# PBIX Model Extractor

A GUI application for extracting Power BI model metadata from .pbix files.

## System Requirements
- Windows 10 or later
- Power BI Desktop (required for pbi-tools operation)
- No additional Python installation needed (bundled with executable)

## ⚠️ Platform Compatibility
This application is **Windows-only**. It is not compatible with Linux or Mac operating systems.

## Installation
1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file to your preferred location
3. Run `pbix-extractor.exe`

## Usage
1. Launch the application
2. Select your PBIX file using the Browse button
3. Choose an output directory
4. Click OK to extract metadata

The application will extract and provide paths to:
- **Primary Model Metadata**: DataModelSchema/model.json
- Additional files:
  - ReportMetadata.json
  - Connections.json
  - DiagramLayout.json

## Development
If you want to run from source:
```bash
python main.py
```

## License
MIT License