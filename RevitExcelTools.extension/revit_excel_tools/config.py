"""
Configuration management for Revit Excel Tools
"""
import os
import json
from System.IO import Path

class Config(object):
    """Manage tool configuration settings"""

    def __init__(self):
        # Get the extension directory
        ext_dir = os.path.dirname(os.path.dirname(__file__))
        self.config_file = os.path.join(ext_dir, "config.json")
        self._config = self._load_config()

    def _load_config(self):
        """Load configuration from JSON file"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading config: {e}")
                return self._get_defaults()
        return self._get_defaults()

    def _get_defaults(self):
        """Get default configuration values"""
        return {
            "excel_file_path": "",
            "export_all_elements": True,
            "include_parameters": True,
            "auto_backup": True,
            "sheet_name": "RevitData",
            "element_filter": ["Walls", "Doors", "Windows", "Floors"],
        }

    def get(self, key, default=None):
        """Get a configuration value"""
        return self._config.get(key, default)

    def set(self, key, value):
        """Set a configuration value"""
        self._config[key] = value
        self._save_config()

    def _save_config(self):
        """Save configuration to JSON file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self._config, f, indent=2)
        except Exception as e:
            print(f"Error saving config: {e}")

    def to_dict(self):
        """Return all configuration as dictionary"""
        return self._config.copy()
