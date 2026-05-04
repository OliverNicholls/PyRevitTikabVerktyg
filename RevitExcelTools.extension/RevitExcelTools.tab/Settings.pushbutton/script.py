"""
Settings and configuration for Revit Excel Tools
"""
from pyrevit import forms
from revit_excel_tools.config import Config

__title__ = "Settings"
__doc__ = "Configure Revit Excel Tools settings"


def main():
    try:
        config = Config()

        # Create a simple form with settings
        settings_dict = {
            "Excel File Path": config.get("excel_file_path", ""),
            "Export All Elements": config.get("export_all_elements", True),
            "Include Parameters": config.get("include_parameters", True),
            "Auto-backup on Import": config.get("auto_backup", True),
        }

        # Display settings info
        msg = "Current Settings:\n\n"
        for key, value in settings_dict.items():
            msg += f"{key}: {value}\n"

        forms.alert(msg, title="Revit Excel Tools Settings")
        print("Settings displayed. Edit config.py to change settings.")

    except Exception as e:
        print(f"Error loading settings: {str(e)}")
        forms.alert(f"Error: {str(e)}", title="Error")


if __name__ == "__main__":
    main()
