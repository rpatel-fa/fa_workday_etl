"""Configuration parser for Excel ingestion."""

from pathlib import Path
from typing import Union

from ..models import ExcelIngestConfig, FileType, WorkbookConfig, SheetConfig
from .base_parser import BaseConfigParser


class IngestConfigParser(BaseConfigParser):
    """Parser for ingest JSON configuration files.

    Supports loading configuration from:
    - File path (str or Path)
    - JSON string
    - Python dict/list

    JSON format:
    [
        {
            "workbookFileName": "/path/to/workbook.xlsx",
            "workbookPathName": "/optional/base/dir",
            "fileType": "EXCEL",
            "sheets": [
                {
                    "sheetName": "Sheet1",
                    "targetTableName": "my_table",
                    "headerRow": 1,
                    "dataRow": 2
                }
            ]
        }
    ]
    """

    @classmethod
    def from_json(
        cls,
        json_data: Union[str, Path, dict, list],
        database_path: Union[str, Path],
        create_tables: bool = True,
        replace_data: bool = True,
    ) -> ExcelIngestConfig:
        """Load configuration from JSON.

        Args:
            json_data: JSON file path, JSON string, dict, or list of workbook configs.
            database_path: Path to the DuckDB database file.
            create_tables: Whether to create tables if they don't exist.
            replace_data: Whether to replace existing data or append.

        Returns:
            ExcelIngestConfig instance.

        Raises:
            FileNotFoundError: If json_data is a path that doesn't exist.
            json.JSONDecodeError: If JSON parsing fails.
            ValueError: If configuration is invalid.
        """
        data = cls.load_json_data(json_data)

        # Parse workbook configurations
        workbooks = cls._parse_workbooks(data)

        # Convert database_path to Path
        db_path = Path(database_path) if isinstance(database_path, str) else database_path

        return ExcelIngestConfig(
            database_path=db_path,
            workbooks=workbooks,
            create_tables=create_tables,
            replace_data=replace_data,
        )

    @staticmethod
    def _parse_workbooks(data: list) -> list[WorkbookConfig]:
        """Parse list of workbook configurations.

        Args:
            data: List of workbook config dictionaries.

        Returns:
            List of WorkbookConfig instances.
        """
        workbooks = []
        for wb_data in data:
            if "workbookFileName" not in wb_data:
                raise ValueError("Missing required field: workbookFileName")

            workbook_path_name = (
                wb_data.get("workbookPathName")
                or wb_data.get("workbookPath")
                or wb_data.get("dataPath")
            )

            workbook_file_name = wb_data["workbookFileName"]

            file_type_raw = wb_data.get("fileType")
            if file_type_raw is None:
                suffix = Path(str(workbook_file_name)).suffix.lower()
                if suffix in {".csv", ".tsv", ".psv", ".txt"}:
                    file_type = FileType.DELIMITED
                else:
                    file_type = FileType.EXCEL
            else:
                file_type_str = str(file_type_raw).upper()
                if file_type_str in {"CSV", "TSV", "PSV", "TXT"}:
                    file_type = FileType.DELIMITED
                elif file_type_str in {"XLS", "XLSX", "XLSM"}:
                    file_type = FileType.EXCEL
                else:
                    try:
                        file_type = FileType(file_type_str)
                    except ValueError:
                        raise ValueError(f"Unsupported fileType: {file_type_str}")

            sheets = IngestConfigParser._parse_sheets(
                wb_data.get("sheets", []),
                allow_missing_sheet_name=(file_type == FileType.DELIMITED),
                default_sheet_name=Path(str(workbook_file_name)).stem,
            )

            delimiter_raw = (
                wb_data.get("delimiter")
                or wb_data.get("sep")
                or wb_data.get("separator")
            )
            delimiter = str(delimiter_raw) if delimiter_raw is not None else None

            workbook = WorkbookConfig(
                workbook_file_name=workbook_file_name,
                workbook_path_name=workbook_path_name,
                sheets=sheets,
                file_type=file_type,
                delimiter=delimiter,
            )
            workbooks.append(workbook)

        return workbooks

    @staticmethod
    def _parse_sheets(
        sheets_data: list,
        *,
        allow_missing_sheet_name: bool = False,
        default_sheet_name: str = "*",
    ) -> list[SheetConfig]:
        """Parse list of sheet configurations.

        Args:
            sheets_data: List of sheet config dictionaries.

        Returns:
            List of SheetConfig instances.
        """
        sheets = []
        for sheet_data in sheets_data:
            if "sheetName" not in sheet_data:
                if allow_missing_sheet_name:
                    sheet_data = dict(sheet_data)
                    sheet_data["sheetName"] = default_sheet_name
                else:
                    raise ValueError("Missing required field: sheetName")
            if "targetTableName" not in sheet_data:
                raise ValueError("Missing required field: targetTableName")

            # Parse headerRow (handle string or int)
            header_row_raw = sheet_data.get("headerRow", 1)
            header_row = int(header_row_raw) if header_row_raw is not None else 1

            # Parse dataRow (handle string or int, None means default)
            data_row_raw = sheet_data.get("dataRow")
            data_row = int(data_row_raw) if data_row_raw is not None else None

            column_headers_raw = (
                sheet_data.get("columnHeaders")
                or sheet_data.get("column_headers")
                or sheet_data.get("columns")
            )
            if column_headers_raw is None:
                column_headers = None
            elif isinstance(column_headers_raw, list):
                column_headers = [str(c) for c in column_headers_raw]
            else:
                raise ValueError(
                    f"columnHeaders must be a list of strings for sheet '{sheet_data.get('sheetName', '')}'"
                )

            sheet = SheetConfig(
                sheet_name=sheet_data["sheetName"],
                target_table_name=sheet_data["targetTableName"],
                header_row=header_row,
                data_row=data_row,
                column_headers=column_headers,
            )
            sheets.append(sheet)

        return sheets

    @classmethod
    def to_json(
        cls,
        config: ExcelIngestConfig,
        filepath: Union[str, Path, None] = None,
    ) -> str:
        """Serialize configuration to JSON.

        Args:
            config: ExcelIngestConfig instance to serialize.
            filepath: Optional file path to write JSON to.

        Returns:
            JSON string representation.
        """
        data = []
        for workbook in config.workbooks:
            wb_data = {
                "workbookFileName": workbook.workbook_file_name,
                "fileType": workbook.file_type.value,
                "sheets": [
                    {
                        "sheetName": sheet.sheet_name,
                        "targetTableName": sheet.target_table_name,
                        "headerRow": sheet.header_row,
                        "dataRow": sheet.data_row,
                        **({"columnHeaders": sheet.column_headers} if sheet.column_headers else {}),
                    }
                    for sheet in workbook.sheets
                ],
            }
            if workbook.workbook_path_name:
                wb_data["workbookPathName"] = workbook.workbook_path_name
            data.append(wb_data)

        json_str = cls.to_json_string(data)

        if filepath:
            cls.write_json_file(data, filepath)

        return json_str

    @staticmethod
    def find_workbook(
        config: ExcelIngestConfig,
        file_name: str,
    ) -> WorkbookConfig | None:
        """Find a workbook configuration by file name.

        Args:
            config: ExcelIngestConfig containing workbook definitions.
            file_name: Workbook file name to match (supports ~ expansion).

        Returns:
            WorkbookConfig if found, None otherwise.
        """
        file_name_expanded = str(Path(file_name).expanduser())

        for workbook in config.workbooks:
            workbook_path_expanded = str(IngestConfigParser._workbook_full_path(workbook))
            if workbook_path_expanded == file_name_expanded or workbook.workbook_file_name == file_name:
                return workbook

        return None

    @staticmethod
    def _workbook_full_path(workbook: WorkbookConfig) -> Path:
        if workbook.workbook_path_name:
            return Path(workbook.workbook_path_name).expanduser() / workbook.workbook_file_name
        return Path(workbook.workbook_file_name).expanduser()

    @staticmethod
    def get_workbooks(
        config: ExcelIngestConfig,
        file_name_filter: str = "*",
    ) -> list[WorkbookConfig]:
        """Get workbooks from config with optional filtering.

        Args:
            config: ExcelIngestConfig containing workbook definitions.
            file_name_filter: Workbook file name to filter on, or "*" for all workbooks.

        Returns:
            List of WorkbookConfig matching the filter.

        Raises:
            ValueError: If a specific file_name_filter is provided but not found.
        """
        if file_name_filter == "*":
            return config.workbooks

        file_name_expanded = str(Path(file_name_filter).expanduser())

        for workbook in config.workbooks:
            workbook_path_expanded = str(IngestConfigParser._workbook_full_path(workbook))
            if workbook_path_expanded == file_name_expanded or workbook.workbook_file_name == file_name_filter:
                return [workbook]

        raise ValueError(f"No workbook config found for: {file_name_filter}")

    @staticmethod
    def get_sheets(
        workbook: WorkbookConfig,
        sheet_filter: str = "*",
    ) -> list[SheetConfig]:
        """Get sheets from a workbook with optional filtering.

        Args:
            workbook: WorkbookConfig containing sheet definitions.
            sheet_filter: Sheet name to filter on, or "*" for all sheets.

        Returns:
            List of SheetConfig matching the filter.
        """
        if sheet_filter == "*":
            return workbook.sheets

        return [
            sheet for sheet in workbook.sheets
            if sheet.sheet_name == sheet_filter
        ]


# Backward compatibility alias
JsonConfigParser = IngestConfigParser
