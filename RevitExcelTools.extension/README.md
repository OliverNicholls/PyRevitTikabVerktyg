# Revit Excel Tools

A PyRevit extension for importing and exporting data between Revit and Excel spreadsheets.

## Features

- **Export to Excel**: Export Revit element data (walls, doors, windows, etc.) to Excel format
- **Import from Excel**: Update Revit elements with data from Excel spreadsheets
- **Configurable**: Customize which elements and parameters to work with
- **Error Handling**: Comprehensive logging and error reporting

## Installation

1. Place the `RevitExcelTools.extension` folder in your PyRevit extensions directory:
   - Windows: `%APPDATA%\pyRevit-Master\extensions\`

2. Restart Revit to load the extension

3. The extension will appear as a new tab in the Revit ribbon named "RevitExcelTools"

## Usage

### Export to Excel

1. Open your Revit project
2. Click the **Export to Excel** button in the RevitExcelTools tab
3. Choose a location and filename for the Excel file
4. The tool will export all selected element data to the Excel file
5. You'll receive a confirmation message with the number of exported elements

### Import from Excel

1. Prepare your Excel file with the following structure:
   - First row: Column headers (must match Revit parameter names)
   - Subsequent rows: Data to import
   - Must include an "ElementId" column to match with Revit elements

2. Open your Revit project with the elements you want to update
3. Click the **Import from Excel** button in the RevitExcelTools tab
4. Select your Excel file
5. The tool will update matching Revit elements with the new data
6. You'll receive a summary of updated, failed, and skipped elements

### Settings

Click the **Settings** button to view current configuration. To modify settings, edit the `config.json` file in the extension folder.

## Configuration

Edit `config.json` to customize behavior:

```json
{
  "excel_file_path": "",
  "export_all_elements": true,
  "include_parameters": true,
  "auto_backup": true,
  "sheet_name": "RevitData",
  "element_filter": [
    "Walls",
    "Doors",
    "Windows",
    "Floors",
    "Columns",
    "Beams"
  ]
}
```

### Configuration Options

- **excel_file_path**: Default path for Excel files (optional)
- **export_all_elements**: If true, export all elements; if false, use element_filter
- **include_parameters**: Include element parameters in export
- **auto_backup**: Automatically backup Revit document before import
- **sheet_name**: Name of the worksheet in Excel file
- **element_filter**: List of element categories to include when export_all_elements is false

## Project Structure

```
RevitExcelTools.extension/
├── manifest.json              # Extension metadata
├── config.json               # Configuration file
├── README.md                 # This file
├── RevitExcelTools.tab/      # Main tab for UI
│   ├── Export.pushbutton/    # Export to Excel button
│   │   └── script.py
│   ├── Import.pushbutton/    # Import from Excel button
│   │   └── script.py
│   └── Settings.pushbutton/  # Settings button
│       └── script.py
└── lib/                      # Utility modules
    ├── __init__.py
    ├── config.py            # Configuration management
    ├── revit_utils.py       # Revit data operations
    └── excel_utils.py       # Excel import/export
```

## Excel Library Support

The tool automatically detects and uses the best available Excel library:

1. **openpyxl** (Recommended) - Modern, pure Python, supports .xlsx
2. **xlrd/xlwt** - Older but reliable, supports .xls
3. **COM (Excel Automation)** - Built-in on Windows, slower but always available

## Development

### Adding Custom Element Types

Edit `lib/revit_utils.py` to customize element data collection:

1. Modify `RevitDataCollector._extract_element_data()` to add custom properties
2. Update the parameter extraction in `_get_element_parameters()`

### Adding Custom Parameters

To work with specific parameters:

1. Edit the element_filter in `config.json`
2. Ensure parameter names match exactly in your Excel files
3. Parameters must not be read-only in Revit

## Troubleshooting

### Import fails with "Element not found"
- Verify the ElementId in your Excel file matches actual Revit elements
- Check that the Revit document containing those elements is open

### Excel file not saving
- Ensure you have write permissions to the selected directory
- Check that Excel is not already open with the file
- Try using a different filename

### Parameters not updating
- Verify parameter names in Excel exactly match Revit parameter names
- Check that parameters are not read-only in Revit
- Some built-in parameters cannot be modified through scripts

### Library import errors
- Install openpyxl: `pip install openpyxl`
- Or install xlrd/xlwt: `pip install xlrd xlwt`
- Falls back to COM interface if libraries unavailable

## Future Enhancements

- [ ] GUI for settings configuration
- [ ] Support for nested parameter updates
- [ ] Batch processing multiple files
- [ ] Parameter mapping/transformation
- [ ] Excel template generation
- [ ] Data validation rules
- [ ] Export/import history tracking

## License

Tikab Strukturmekanik AB

## Support

For issues or feature requests, contact the development team.
