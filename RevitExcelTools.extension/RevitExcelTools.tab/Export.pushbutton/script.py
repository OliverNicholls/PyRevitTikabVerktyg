"""
Export Revit data to Excel
"""
from pyrevit import revit, forms
from revit_excel_tools.excel_utils import ExcelExporter
from revit_excel_tools.revit_utils import RevitDataCollector

__title__ = "Export to Excel"
__doc__ = "Export Revit element data to an Excel spreadsheet"


def main():
    try:
        # Get the active document
        doc = revit.active.document

        # Show file dialog to select output Excel file
        file_path = forms.save_file(
            file_ext="xlsx",
            default_name="revit_export.xlsx"
        )

        if not file_path:
            forms.alert("Export cancelled.")
            return

        # Collect data from Revit
        print("Collecting data from Revit...")
        collector = RevitDataCollector(doc)
        data = collector.collect_elements()

        # Export to Excel
        print(f"Exporting to {file_path}...")
        exporter = ExcelExporter(file_path)
        exporter.export_data(data)

        forms.alert(
            f"Successfully exported {len(data)} elements to Excel.",
            title="Export Complete"
        )
        print("Export completed successfully!")

    except Exception as e:
        print(f"Error during export: {str(e)}")
        forms.alert(f"Export failed: {str(e)}", title="Error")


if __name__ == "__main__":
    main()
