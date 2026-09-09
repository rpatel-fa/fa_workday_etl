from pathlib import Path
from typing import TYPE_CHECKING

from ..models import TransformResult
from .db import open_connection
from .order_reader import OrderReader
from .sql_file_executor import SqlFileExecutor

if TYPE_CHECKING:
    from ..reporting import PipelineReporter


class SqlExecutor:
    def __init__(
        self,
        transform_path: Path,
        database_path: Path,
        reporter: "PipelineReporter | None" = None,
    ):
        self.transform_path = Path(transform_path)
        self.database_path = Path(database_path)
        self.reporter = reporter

    def execute(self) -> list[TransformResult]:
        reader = OrderReader(self.transform_path)
        if not reader.exists():
            return []
        sql_files = reader.read()
        results: list[TransformResult] = []
        with open_connection(self.database_path) as conn:
            runner = SqlFileExecutor(self.transform_path, self.reporter)
            for sql_file in sql_files:
                result = runner.run(conn, sql_file)
                results.append(result)
                if not result.success:
                    if self.reporter:
                        self.reporter.print_transform_abort_on_failure()
                    break
        return results

    def get_sql_file_count(self) -> int:
        reader = OrderReader(self.transform_path)
        if not reader.exists():
            return 0
        return len(reader.read())
