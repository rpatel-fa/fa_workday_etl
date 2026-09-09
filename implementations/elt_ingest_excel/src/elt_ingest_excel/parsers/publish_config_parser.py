"""Configuration parser for publish/output configuration."""

from pathlib import Path
from typing import Union

from ..models import PublishConfig, PublishWorkbookConfig, PublishSheetConfig
from .base_parser import BaseConfigParser


class PublishConfigParser(BaseConfigParser):
    """Parser for publish JSON configuration files.

    JSON format:
    [
        {
            "srcWorkbookPathName": "/path/to/templates",
            "srcWorkbookFileName": "template.xlsm",
            "tgtWorkbookPathName": "/path/to/output",
            "tgtWorkbookFileName": "output_name",
            "sheets": [
                {
                    "srcTableName": "table_name",
                    "sheetName": "Sheet Name",
                    "headerRow": "3",
                    "dataRow": "4"
                }
            ]
        }
    ]
    """

    @classmethod
    def from_json(
        cls,
        json_data: Union[str, Path, dict, list],
    ) -> PublishConfig:
        """Load publish configuration from JSON.

        Args:
            json_data: JSON file path, JSON string, dict, or list of workbook configs.

        Returns:
            PublishConfig instance.

        Raises:
            FileNotFoundError: If json_data is a path that doesn't exist.
            json.JSONDecodeError: If JSON parsing fails.
            ValueError: If configuration is invalid.
        """
        data = cls.load_json_data(json_data)

        # Parse workbook configurations
        workbooks = cls._parse_workbooks(data)

        return PublishConfig(workbooks=workbooks)

    @staticmethod
    def _parse_workbooks(data: list) -> list[PublishWorkbookConfig]:
        """Parse list of publish workbook configurations.

        Args:
            data: List of workbook config dictionaries.

        Returns:
            List of PublishWorkbookConfig instances.
        """
        workbooks = []
        for wb_data in data:
            required_fields = [
                "srcWorkbookPathName",
                "srcWorkbookFileName",
                "tgtWorkbookPathName",
                "tgtWorkbookFileName",
            ]
            for field in required_fields:
                if field not in wb_data:
                    raise ValueError(f"Missing required field: {field}")

            sheets = PublishConfigParser._parse_sheets(wb_data.get("sheets", []))

            workbook = PublishWorkbookConfig(
                src_workbook_path=wb_data["srcWorkbookPathName"],
                src_workbook_file_name=wb_data["srcWorkbookFileName"],
                tgt_workbook_path=wb_data["tgtWorkbookPathName"],
                tgt_workbook_file_name=wb_data["tgtWorkbookFileName"],
                sheets=sheets,
            )
            workbooks.append(workbook)

        return workbooks

    @staticmethod
    def _parse_sheets(sheets_data: list) -> list[PublishSheetConfig]:
        """Parse list of publish sheet configurations.

        Args:
            sheets_data: List of sheet config dictionaries.

        Returns:
            List of PublishSheetConfig instances.
        """
        sheets = []
        for sheet_data in sheets_data:
            if "srcTableName" not in sheet_data:
                raise ValueError("Missing required field: srcTableName")
            if "sheetName" not in sheet_data:
                raise ValueError("Missing required field: sheetName")

            # Parse headerRow (handle string or int)
            header_row_raw = sheet_data.get("headerRow", 1)
            header_row = int(header_row_raw) if header_row_raw is not None else 1

            # Parse dataRow (handle string or int)
            data_row_raw = sheet_data.get("dataRow", 2)
            data_row = int(data_row_raw) if data_row_raw is not None else 2

            use_column_name_index_raw = sheet_data.get(
                "use_column_name_index",
                sheet_data.get("useColumnNameIndex", False),
            )
            if isinstance(use_column_name_index_raw, bool):
                use_column_name_index = use_column_name_index_raw
            elif use_column_name_index_raw is None:
                use_column_name_index = False
            elif isinstance(use_column_name_index_raw, (int, float)):
                use_column_name_index = bool(use_column_name_index_raw)
            else:
                normalized = str(use_column_name_index_raw).strip().upper()
                truthy = {"Y", "YES", "TRUE", "T", "1"}
                falsy = {"N", "NO", "FALSE", "F", "0", ""}
                if normalized in truthy:
                    use_column_name_index = True
                elif normalized in falsy:
                    use_column_name_index = False
                else:
                    raise ValueError(
                        f"Invalid value for use_column_name_index: {use_column_name_index_raw!r}"
                    )

            sheet = PublishSheetConfig(
                src_table_name=sheet_data["srcTableName"],
                sheet_name=sheet_data["sheetName"],
                header_row=header_row,
                data_row=data_row,
                use_column_name_index=use_column_name_index,
            )
            sheets.append(sheet)

        return sheets
