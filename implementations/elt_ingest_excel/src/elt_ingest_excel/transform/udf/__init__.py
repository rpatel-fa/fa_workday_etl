from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import duckdb


def register_all(conn: "duckdb.DuckDBPyConnection") -> None:
    from . import address
    from . import phone

    address.register(conn)
    phone.register(conn)
