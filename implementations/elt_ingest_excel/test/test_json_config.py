"""Tests for JSON configuration parsing."""

import json
import tempfile
from pathlib import Path

import pytest

from elt_ingest_excel import (
    JsonConfigParser,
    ExcelIngestConfig,
    FileType,
    WorkbookConfig,
    SheetConfig,
)


# Path to test fixtures
FIXTURES_DIR = Path(__file__).parent / "fixtures"


class TestJsonConfigParser:
    """Tests for JsonConfigParser."""

    def test_from_json_file_single_workbook(self):
        """Test loading config from a JSON file with single workbook."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "single_workbook.json",
            database_path="/path/to/db.duckdb",
        )

        assert isinstance(config, ExcelIngestConfig)
        assert config.database_path == Path("/path/to/db.duckdb")
        assert len(config.workbooks) == 1
        assert config.workbooks[0].workbook_file_name == "/path/to/workbook.xlsx"
        assert len(config.workbooks[0].sheets) == 1
        assert config.workbooks[0].sheets[0].sheet_name == "Sheet1"
        assert config.workbooks[0].sheets[0].target_table_name == "table1"

    def test_from_json_file_multiple_workbooks(self):
        """Test loading config from a JSON file with multiple workbooks."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "multiple_workbooks.json",
            database_path="/path/to/db.duckdb",
        )

        assert len(config.workbooks) == 2
        assert config.workbooks[0].workbook_file_name == "/path/to/workbook1.xlsx"
        assert config.workbooks[1].workbook_file_name == "/path/to/workbook2.xlsx"

    def test_from_json_string_path(self):
        """Test loading config using string path."""
        config = JsonConfigParser.from_json(
            str(FIXTURES_DIR / "single_workbook.json"),
            database_path="/path/to/db.duckdb",
        )

        assert len(config.workbooks) == 1
        assert config.workbooks[0].workbook_file_name == "/path/to/workbook.xlsx"

    def test_sheet_optional_fields_defaults(self):
        """Test that sheet optional fields have correct defaults."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "single_workbook.json",
            database_path="/path/to/db.duckdb",
        )

        sheet = config.workbooks[0].sheets[0]
        assert sheet.header_row == 1
        assert sheet.data_row == 2  # Default is header_row + 1

    def test_sheet_custom_header_row(self):
        """Test loading sheet with custom header row settings."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "custom_header_row.json",
            database_path="/path/to/db.duckdb",
        )

        sheet = config.workbooks[0].sheets[0]
        assert sheet.header_row == 3
        assert sheet.data_row == 4

    def test_sheet_headerless_with_column_headers(self):
        config = JsonConfigParser.from_json(
            [
                {
                    "workbookFileName": "/path/to/workbook.xlsx",
                    "fileType": "EXCEL",
                    "sheets": [
                        {
                            "sheetName": "Sheet1",
                            "targetTableName": "table1",
                            "headerRow": 0,
                            "dataRow": 11,
                            "columnHeaders": ["spreadsheet_key", "error_description"],
                        }
                    ],
                }
            ],
            database_path="/path/to/db.duckdb",
        )

        sheet = config.workbooks[0].sheets[0]
        assert sheet.header_row == 0
        assert sheet.data_row == 11
        assert sheet.column_headers == ["spreadsheet_key", "error_description"]

    def test_delimited_infers_file_type_and_allows_missing_sheet_name(self):
        config = JsonConfigParser.from_json(
            [
                {
                    "workbookFileName": "/path/to/data.csv",
                    "sheets": [
                        {
                            "targetTableName": "csv_table",
                            "headerRow": 1,
                            "dataRow": 2,
                        }
                    ],
                }
            ],
            database_path="/path/to/db.duckdb",
        )

        assert config.workbooks[0].file_type == FileType.DELIMITED
        assert config.workbooks[0].sheets[0].sheet_name == "data"

    def test_config_defaults(self):
        """Test that config has correct default values."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "single_workbook.json",
            database_path="/path/to/db.duckdb",
        )

        assert config.create_tables is True
        assert config.replace_data is True

    def test_config_custom_flags(self):
        """Test setting custom create_tables and replace_data flags."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "single_workbook.json",
            database_path="/path/to/db.duckdb",
            create_tables=False,
            replace_data=False,
        )

        assert config.create_tables is False
        assert config.replace_data is False

    def test_to_json(self):
        """Test serializing config to JSON."""
        config = ExcelIngestConfig(
            database_path=Path("/path/to/db.duckdb"),
            workbooks=[
                WorkbookConfig(
                    workbook_file_name="/path/to/workbook.xlsx",
                    sheets=[
                        SheetConfig(
                            sheet_name="Sheet1",
                            target_table_name="table1",
                            header_row=2,
                            data_row=3,
                        )
                    ],
                )
            ],
        )

        json_str = JsonConfigParser.to_json(config)
        data = json.loads(json_str)

        assert len(data) == 1
        assert data[0]["workbookFileName"] == "/path/to/workbook.xlsx"
        assert data[0]["sheets"][0]["sheetName"] == "Sheet1"
        assert data[0]["sheets"][0]["targetTableName"] == "table1"
        assert data[0]["sheets"][0]["headerRow"] == 2
        assert data[0]["sheets"][0]["dataRow"] == 3

    def test_to_json_includes_column_headers_when_present(self):
        config = ExcelIngestConfig(
            database_path=Path("/path/to/db.duckdb"),
            workbooks=[
                WorkbookConfig(
                    workbook_file_name="/path/to/workbook.xlsx",
                    sheets=[
                        SheetConfig(
                            sheet_name="Sheet1",
                            target_table_name="table1",
                            header_row=0,
                            data_row=11,
                            column_headers=["spreadsheet_key", "error_description"],
                        )
                    ],
                )
            ],
        )

        json_str = JsonConfigParser.to_json(config)
        data = json.loads(json_str)
        assert data[0]["sheets"][0]["columnHeaders"] == [
            "spreadsheet_key",
            "error_description",
        ]

    def test_to_json_file(self):
        """Test serializing config to a JSON file."""
        config = ExcelIngestConfig(
            database_path=Path("/path/to/db.duckdb"),
            workbooks=[
                WorkbookConfig(
                    workbook_file_name="/path/to/workbook.xlsx",
                    sheets=[
                        SheetConfig(
                            sheet_name="Sheet1",
                            target_table_name="table1",
                        )
                    ],
                )
            ],
        )

        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            temp_path = f.name

        try:
            JsonConfigParser.to_json(config, filepath=temp_path)

            # Read back and verify
            with open(temp_path) as f:
                data = json.load(f)

            assert len(data) == 1
            assert data[0]["workbookFileName"] == "/path/to/workbook.xlsx"
        finally:
            Path(temp_path).unlink()

    def test_roundtrip_json(self):
        """Test that config can be serialized and deserialized."""
        # Load original config
        original_config = JsonConfigParser.from_json(
            FIXTURES_DIR / "custom_header_row.json",
            database_path="/path/to/db.duckdb",
        )

        # Serialize to temp file
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            temp_path = f.name

        try:
            JsonConfigParser.to_json(original_config, filepath=temp_path)

            # Deserialize back
            reloaded_config = JsonConfigParser.from_json(
                temp_path,
                database_path="/path/to/db.duckdb",
            )

            # Verify roundtrip
            assert len(reloaded_config.workbooks) == len(original_config.workbooks)
            assert (
                reloaded_config.workbooks[0].workbook_file_name
                == original_config.workbooks[0].workbook_file_name
            )
            assert (
                reloaded_config.workbooks[0].sheets[0].header_row
                == original_config.workbooks[0].sheets[0].header_row
            )
        finally:
            Path(temp_path).unlink()

    def test_missing_workbook_filename(self):
        """Test error when workbookFileName is missing."""
        with pytest.raises(ValueError, match="Missing required field: workbookFileName"):
            JsonConfigParser.from_json(
                FIXTURES_DIR / "missing_workbook_filename.json",
                database_path="/path/to/db.duckdb",
            )

    def test_missing_sheet_name(self):
        """Test error when sheetName is missing."""
        with pytest.raises(ValueError, match="Missing required field: sheetName"):
            JsonConfigParser.from_json(
                FIXTURES_DIR / "missing_sheet_name.json",
                database_path="/path/to/db.duckdb",
            )

    def test_missing_target_table_name(self):
        """Test error when targetTableName is missing."""
        with pytest.raises(ValueError, match="Missing required field: targetTableName"):
            JsonConfigParser.from_json(
                FIXTURES_DIR / "missing_target_table_name.json",
                database_path="/path/to/db.duckdb",
            )

    def test_file_not_found(self):
        """Test error when config file doesn't exist."""
        with pytest.raises(FileNotFoundError):
            JsonConfigParser.from_json(
                Path("/nonexistent/config.json"),
                database_path="/path/to/db.duckdb",
            )

    def test_invalid_json_file(self):
        """Test error when JSON file contains invalid JSON."""
        with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as f:
            f.write("not valid json")
            temp_path = f.name

        try:
            with pytest.raises(json.JSONDecodeError):
                JsonConfigParser.from_json(
                    temp_path,
                    database_path="/path/to/db.duckdb",
                )
        finally:
            Path(temp_path).unlink()


class TestFindWorkbookAndGetSheets:
    """Tests for find_workbook and get_sheets helper methods."""

    @pytest.fixture
    def config_with_multiple_sheets(self):
        """Load config with multiple workbooks and sheets."""
        return JsonConfigParser.from_json(
            FIXTURES_DIR / "workbook_multiple_sheets.json",
            database_path="/path/to/db.duckdb",
        )

    def test_find_workbook_exact_match(self, config_with_multiple_sheets):
        """Test finding workbook by exact file path."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        assert workbook is not None
        assert workbook.workbook_file_name == "/data/sales.xlsx"
        assert len(workbook.sheets) == 3

    def test_find_workbook_with_tilde_expansion(self, config_with_multiple_sheets):
        """Test finding workbook with ~ path expansion."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "~/Documents/finance.xlsx",
        )

        assert workbook is not None
        assert workbook.workbook_file_name == "~/Documents/finance.xlsx"
        assert len(workbook.sheets) == 2

    def test_find_workbook_not_found(self, config_with_multiple_sheets):
        """Test finding workbook that doesn't exist returns None."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/nonexistent/workbook.xlsx",
        )

        assert workbook is None

    def test_get_sheets_all(self, config_with_multiple_sheets):
        """Test getting all sheets with '*' filter."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        sheets = JsonConfigParser.get_sheets(workbook, "*")

        assert len(sheets) == 3
        assert sheets[0].sheet_name == "Q1"
        assert sheets[1].sheet_name == "Q2"
        assert sheets[2].sheet_name == "Q3"

    def test_get_sheets_default_filter(self, config_with_multiple_sheets):
        """Test getting all sheets with default filter (no argument)."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        sheets = JsonConfigParser.get_sheets(workbook)

        assert len(sheets) == 3

    def test_get_sheets_specific_sheet(self, config_with_multiple_sheets):
        """Test filtering to a specific sheet by name."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        sheets = JsonConfigParser.get_sheets(workbook, "Q2")

        assert len(sheets) == 1
        assert sheets[0].sheet_name == "Q2"
        assert sheets[0].target_table_name == "sales_q2"

    def test_get_sheets_no_match(self, config_with_multiple_sheets):
        """Test filtering with sheet name that doesn't exist."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        sheets = JsonConfigParser.get_sheets(workbook, "NonexistentSheet")

        assert len(sheets) == 0

    def test_iterate_workbooks_and_sheets(self, config_with_multiple_sheets):
        """Test iterating over config using find_workbook and get_sheets."""
        # Simulate test_read_excel.py pattern
        file_names = ["/data/sales.xlsx", "~/Documents/finance.xlsx"]
        results = []

        for file_name in file_names:
            workbook = JsonConfigParser.find_workbook(
                config_with_multiple_sheets,
                file_name,
            )
            assert workbook is not None

            for sheet in JsonConfigParser.get_sheets(workbook, "*"):
                results.append({
                    "workbook": workbook.workbook_file_name,
                    "sheet": sheet.sheet_name,
                    "table": sheet.target_table_name,
                })

        assert len(results) == 5
        assert results[0]["sheet"] == "Q1"
        assert results[3]["sheet"] == "Revenue"
        assert results[4]["table"] == "finance_expenses"

    def test_iterate_with_sheet_filter(self, config_with_multiple_sheets):
        """Test iterating with specific sheet filter."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        # Only process Q1 and Q3
        results = []
        for sheet_filter in ["Q1", "Q3"]:
            for sheet in JsonConfigParser.get_sheets(workbook, sheet_filter):
                results.append(sheet.sheet_name)

        assert results == ["Q1", "Q3"]

    def test_sheet_config_preserved(self, config_with_multiple_sheets):
        """Test that sheet config values are preserved through iteration."""
        workbook = JsonConfigParser.find_workbook(
            config_with_multiple_sheets,
            "/data/sales.xlsx",
        )

        sheets = JsonConfigParser.get_sheets(workbook, "Q3")

        assert len(sheets) == 1
        sheet = sheets[0]
        assert sheet.header_row == 2
        assert sheet.data_row == 3
        assert sheet.target_table_name == "sales_q3"


class TestFileType:
    """Tests for FileType parsing and handling."""

    def test_file_type_default_is_excel(self):
        """Test that file_type defaults to EXCEL when not specified."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "single_workbook.json",
            database_path="/path/to/db.duckdb",
        )

        assert config.workbooks[0].file_type == FileType.EXCEL

    def test_file_type_explicit_excel(self):
        """Test parsing explicit EXCEL fileType."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "with_file_type.json",
            database_path="/path/to/db.duckdb",
        )

        assert config.workbooks[0].file_type == FileType.EXCEL
        assert config.workbooks[0].workbook_file_name == "/path/to/workbook.xlsx"

    def test_file_type_delimited(self):
        """Test parsing DELIMITED fileType."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "with_file_type.json",
            database_path="/path/to/db.duckdb",
        )

        assert config.workbooks[1].file_type == FileType.DELIMITED
        assert config.workbooks[1].workbook_file_name == "/path/to/data.csv"

    def test_file_type_case_insensitive(self):
        """Test that fileType parsing is case insensitive."""
        json_data = [{
            "workbookFileName": "/path/to/file.xlsx",
            "fileType": "excel",
            "sheets": [{"sheetName": "Sheet1", "targetTableName": "table1"}],
        }]

        config = JsonConfigParser.from_json(
            json_data,
            database_path="/path/to/db.duckdb",
        )

        assert config.workbooks[0].file_type == FileType.EXCEL

    def test_file_type_invalid(self):
        """Test error for invalid fileType."""
        json_data = [{
            "workbookFileName": "/path/to/file.xlsx",
            "fileType": "INVALID_TYPE",
            "sheets": [{"sheetName": "Sheet1", "targetTableName": "table1"}],
        }]

        with pytest.raises(ValueError, match="Unsupported fileType"):
            JsonConfigParser.from_json(
                json_data,
                database_path="/path/to/db.duckdb",
            )

    def test_file_type_in_to_json(self):
        """Test that fileType is included in serialized JSON."""
        config = ExcelIngestConfig(
            database_path=Path("/path/to/db.duckdb"),
            workbooks=[
                WorkbookConfig(
                    workbook_file_name="/path/to/workbook.xlsx",
                    file_type=FileType.EXCEL,
                    sheets=[
                        SheetConfig(
                            sheet_name="Sheet1",
                            target_table_name="table1",
                        )
                    ],
                )
            ],
        )

        json_str = JsonConfigParser.to_json(config)
        data = json.loads(json_str)

        assert data[0]["fileType"] == "EXCEL"

    def test_file_type_roundtrip(self):
        """Test fileType survives JSON roundtrip."""
        config = JsonConfigParser.from_json(
            FIXTURES_DIR / "with_file_type.json",
            database_path="/path/to/db.duckdb",
        )

        json_str = JsonConfigParser.to_json(config)
        data = json.loads(json_str)

        assert data[0]["fileType"] == "EXCEL"
        assert data[1]["fileType"] == "DELIMITED"
