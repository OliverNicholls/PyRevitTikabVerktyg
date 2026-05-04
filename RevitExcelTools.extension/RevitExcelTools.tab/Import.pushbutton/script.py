"""
Import data from Excel into Revit
"""
from pyrevit import revit, forms
from revit_excel_tools.excel_utils import ExcelImporter
from revit_excel_tools.revit_utils import RevitDataUpdater

__title__ = "Import from Excel"
__doc__ = "Import element data from Excel and update Revit elements"


def main():
    try:
        # Get the active document
        doc = revit.active.document

        # Show file dialog to select Excel file
        file_path = forms.pick_file(
            file_ext="xlsx",
            title="Select Excel File to Import"
        )

        if not file_path:
            forms.alert("Import cancelled.")
            return

        # Import data from Excel
        print(f"Importing from {file_path}...")
        importer = ExcelImporter(file_path)
        data = importer.import_data()

        if not data:
            forms.alert("No data found in Excel file.")
            return

        # Update Revit elements
        print(f"Updating {len(data)} elements in Revit...")
        updater = RevitDataUpdater(doc)
        results = updater.update_elements(data)

        # Show results
        forms.alert(
            f"Import complete:\n"
            f"Updated: {results['updated']}\n"
            f"Failed: {results['failed']}\n"
            f"Skipped: {results['skipped']}",
            title="Import Complete"
        )
        print("Import completed successfully!")

    except Exception as e:
        print(f"Error during import: {str(e)}")
        forms.alert(f"Import failed: {str(e)}", title="Error")


if __name__ == "__main__":
    main()
