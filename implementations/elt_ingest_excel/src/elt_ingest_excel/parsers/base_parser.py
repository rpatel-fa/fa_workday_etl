"""Base parser with common JSON loading functionality."""

import json
from pathlib import Path
from typing import Union


class BaseConfigParser:
    """Base class for JSON configuration parsers.

    Provides common functionality for loading JSON from various sources:
    - File path (str or Path)
    - JSON string
    - Python dict/list
    """

    @staticmethod
    def load_json_data(json_data: Union[str, Path, dict, list]) -> list:
        """Load and normalize JSON data from various sources.

        Args:
            json_data: JSON file path, JSON string, dict, or list.

        Returns:
            List of configuration dictionaries.

        Raises:
            FileNotFoundError: If json_data is a path that doesn't exist.
            json.JSONDecodeError: If JSON parsing fails.
            ValueError: If json_data type is not supported.
        """
        # Parse JSON data based on type
        if isinstance(json_data, Path):
            if not json_data.exists():
                raise FileNotFoundError(f"Config file not found: {json_data}")
            data = json.loads(json_data.read_text())
        elif isinstance(json_data, str):
            # Check if it's a file path or JSON string
            path = Path(json_data)
            if path.exists():
                data = json.loads(path.read_text())
            else:
                data = json.loads(json_data)
        elif isinstance(json_data, (dict, list)):
            data = json_data
        else:
            raise ValueError(f"Unsupported json_data type: {type(json_data)}")

        # Ensure data is a list
        if isinstance(data, dict):
            data = [data]

        return data

    @staticmethod
    def write_json_file(data: list, filepath: Union[str, Path]) -> None:
        """Write JSON data to a file.

        Args:
            data: List of configuration dictionaries.
            filepath: Path to write JSON to.
        """
        path = Path(filepath) if isinstance(filepath, str) else filepath
        json_str = json.dumps(data, indent=2)
        path.write_text(json_str)

    @staticmethod
    def to_json_string(data: list) -> str:
        """Convert data to JSON string.

        Args:
            data: List of configuration dictionaries.

        Returns:
            JSON string representation.
        """
        return json.dumps(data, indent=2)
