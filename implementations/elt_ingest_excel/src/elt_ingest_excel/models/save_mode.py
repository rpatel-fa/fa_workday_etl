"""Save mode definitions for database writers."""

from enum import Enum


class SaveMode(Enum):
    """Mode for saving data to a database table.

    Attributes:
        DROP: Drop the table if it exists. Does not recreate or populate.
        RECREATE: Drop the table, recreate based on schema, populate with data.
        OVERWRITE: Delete all rows from existing table, then insert new data.
        APPEND: Add rows to existing table without deleting existing data.
    """
    DROP = "DROP"
    RECREATE = "RECREATE"
    OVERWRITE = "OVERWRITE"
    APPEND = "APPEND"
