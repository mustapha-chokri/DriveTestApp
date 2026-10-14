"""Configuration management for DriveTestApp."""
import os
import json
from pathlib import Path

DEFAULT_CONFIG_FILE = "config.json"

class AppConfig:
    """Manages application settings persisted in a JSON file."""
    
    def __init__(self, filepath=DEFAULT_CONFIG_FILE):
        self.filepath = filepath
        self.rnp_path = ""
        self.adt_path = ""
        self.output_dir = str(Path.home() / "Desktop")
        self.city = ""
        self.load()

    def load(self):
        """Loads configuration from disk if available."""
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.rnp_path = data.get("rnp", self.rnp_path)
                    self.adt_path = data.get("adt", self.adt_path)
                    self.output_dir = data.get("out", self.output_dir)
                    self.city = data.get("city", self.city)
            except Exception:
                # If corrupt or unreadable, fall back to current defaults
                pass

    def save(self, rnp_path=None, adt_path=None, output_dir=None, city=None):
        """Saves current or updated configuration values to disk."""
        if rnp_path is not None:
            self.rnp_path = rnp_path
        if adt_path is not None:
            self.adt_path = adt_path
        if output_dir is not None:
            self.output_dir = output_dir
        if city is not None:
            self.city = city

        data = {
            "rnp": self.rnp_path,
            "adt": self.adt_path,
            "out": self.output_dir,
            "city": self.city
        }
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            raise IOError(f"Failed to save configuration: {e}")
