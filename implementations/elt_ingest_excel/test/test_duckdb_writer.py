"""Tests for DuckDB writer functionality."""

import tempfile
from pathlib import Path

import pandas as pd
import pytest
import duckdb

from elt_ingest_excel import SaveMode, DuckDBWriter, WriteResult


class TestSaveMode:
    """Tests for SaveMode enum."""

    def test_save_mode_values(self):
        """Test that all expected save modes exist."""
        assert SaveMode.DROP.value == "DROP"
        assert SaveMode.RECREATE.value == "RECREATE"
        assert SaveMode.OVERWRITE.value == "OVERWRITE"
        assert SaveMode.APPEND.value == "APPEND"

    def test_save_mode_from_string(self):
        """Test creating SaveMode from string value."""
        assert SaveMode("DROP") == SaveMode.DROP
        assert SaveMode("RECREATE") == SaveMode.RECREATE
        assert SaveMode("OVERWRITE") == SaveMode.OVERWRITE
        assert SaveMode("APPEND") == SaveMode.APPEND


class TestDuckDBWriter:
    """Tests for DuckDBWriter class."""

    @pytest.fixture
    def temp_db(self):
        """Create a temporary DuckDB database."""
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "test.duckdb"
            yield db_path

    @pytest.fixture
    def sample_df(self):
        """Create a sample DataFrame for testing."""
        # Use default pandas object dtype (not explicit str) for DuckDB compatibility
        return pd.DataFrame({
            "id": [1, 2, 3],
            "name": ["Alice", "Bob", "Charlie"],
            "value": [100, 200, 300],
        })

    def test_write_recreate_creates_table(self, temp_db, sample_df):
        """Test RECREATE mode creates a new table."""
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.RECREATE)

        assert result.table_name == "test_table"
        assert result.rows_written == 3
        assert result.row_count == 3
        assert result.save_mode == SaveMode.RECREATE

        # Verify data in database
        with duckdb.connect(str(temp_db)) as conn:
            count = conn.execute("SELECT COUNT(*) FROM test_table").fetchone()[0]
            assert count == 3

    def test_write_recreate_replaces_existing(self, temp_db, sample_df):
        """Test RECREATE mode replaces existing table."""
        # First write
        with DuckDBWriter(temp_db) as writer:
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

        # Second write with different data
        new_df = pd.DataFrame({
            "id": [4, 5],
            "name": ["Dave", "Eve"],
            "value": [400, 500],
        })

        with DuckDBWriter(temp_db) as writer:
            result = writer.write(new_df, "test_table", SaveMode.RECREATE)

        assert result.rows_written == 2
        assert result.row_count == 2

    def test_write_drop_removes_table(self, temp_db, sample_df):
        """Test DROP mode removes the table."""
        # First create table
        with DuckDBWriter(temp_db) as writer:
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

        # Then drop it
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.DROP)

        assert result.rows_written == 0
        assert result.row_count == 0

        # Verify table doesn't exist
        with duckdb.connect(str(temp_db)) as conn:
            tables = conn.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_name = 'test_table'"
            ).fetchall()
            assert len(tables) == 0

    def test_write_overwrite_deletes_and_inserts(self, temp_db, sample_df):
        """Test OVERWRITE mode deletes existing data and inserts new."""
        # First write
        with DuckDBWriter(temp_db) as writer:
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

        # Overwrite with different data
        new_df = pd.DataFrame({
            "id": [4, 5],
            "name": ["Dave", "Eve"],
            "value": [400, 500],
        })

        with DuckDBWriter(temp_db) as writer:
            result = writer.write(new_df, "test_table", SaveMode.OVERWRITE)

        assert result.rows_written == 2
        assert result.row_count == 2

        # Verify only new data exists
        with duckdb.connect(str(temp_db)) as conn:
            names = conn.execute("SELECT name FROM test_table ORDER BY name").fetchall()
            assert [n[0] for n in names] == ["Dave", "Eve"]

    def test_write_overwrite_creates_if_not_exists(self, temp_db, sample_df):
        """Test OVERWRITE mode creates table if it doesn't exist."""
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.OVERWRITE)

        assert result.rows_written == 3
        assert result.row_count == 3

    def test_write_append_adds_rows(self, temp_db, sample_df):
        """Test APPEND mode adds rows to existing table."""
        # First write
        with DuckDBWriter(temp_db) as writer:
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

        # Append more data
        new_df = pd.DataFrame({
            "id": [4, 5],
            "name": ["Dave", "Eve"],
            "value": [400, 500],
        })

        with DuckDBWriter(temp_db) as writer:
            result = writer.write(new_df, "test_table", SaveMode.APPEND)

        assert result.rows_written == 2
        assert result.row_count == 5  # 3 original + 2 new

    def test_write_append_creates_if_not_exists(self, temp_db, sample_df):
        """Test APPEND mode creates table if it doesn't exist."""
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.APPEND)

        assert result.rows_written == 3
        assert result.row_count == 3

    def test_write_empty_dataframe_drop(self, temp_db, sample_df):
        """Test writing an empty DataFrame with DROP mode removes table."""
        # First create table with data
        with DuckDBWriter(temp_db) as writer:
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

        # Then use DROP to remove it
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.DROP)

        assert result.rows_written == 0
        assert result.row_count == 0

    def test_context_manager_required(self, temp_db, sample_df):
        """Test that DuckDBWriter requires context manager."""
        writer = DuckDBWriter(temp_db)

        with pytest.raises(RuntimeError, match="must be used as a context manager"):
            writer.write(sample_df, "test_table", SaveMode.RECREATE)

    def test_write_result_dataclass(self, temp_db, sample_df):
        """Test WriteResult contains all expected fields."""
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table", SaveMode.RECREATE)

        assert isinstance(result, WriteResult)
        assert hasattr(result, "table_name")
        assert hasattr(result, "rows_written")
        assert hasattr(result, "row_count")
        assert hasattr(result, "save_mode")

    def test_table_with_special_characters(self, temp_db, sample_df):
        """Test writing to table with special characters in name."""
        with DuckDBWriter(temp_db) as writer:
            result = writer.write(sample_df, "test_table_123", SaveMode.RECREATE)

        assert result.table_name == "test_table_123"
        assert result.row_count == 3

    def test_path_expansion(self, sample_df):
        """Test that ~ in path is expanded."""
        # Create temp dir in home
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "test.duckdb"

            with DuckDBWriter(db_path) as writer:
                result = writer.write(sample_df, "test_table", SaveMode.RECREATE)

            assert result.row_count == 3
