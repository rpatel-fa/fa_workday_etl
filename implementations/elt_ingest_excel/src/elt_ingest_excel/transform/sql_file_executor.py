from pathlib import Path
from typing import TYPE_CHECKING

import duckdb

from ..models import TransformResult

if TYPE_CHECKING:
    from ..reporting import PipelineReporter


class SqlFileExecutor:
    def __init__(self, transform_path: Path, reporter: "PipelineReporter | None" = None):
        self.transform_path = Path(transform_path).resolve()
        shared_ref = self.transform_path.parent.parent / "ref"
        shared_macro = self.transform_path.parent.parent / "macro"
        self.allowed_paths: list[Path] = [self.transform_path]
        if shared_ref.exists():
            self.allowed_paths.append(shared_ref.resolve())
        if shared_macro.exists():
            self.allowed_paths.append(shared_macro.resolve())
        self.reporter = reporter

    def run(
        self,
        conn: duckdb.DuckDBPyConnection,
        sql_file: str,
    ) -> TransformResult:
        sql_path = (self.transform_path / sql_file).resolve()
        if self.reporter:
            self.reporter.print_sql_file_start(sql_file)

        # Only execute SQL files that live within the configured transform path or shared ref path.
        if not any(sql_path.is_relative_to(p) for p in self.allowed_paths):
            error = f"SQL file '{sql_file}' is outside the allowed config path and will not be executed"
            if self.reporter:
                self.reporter.print_sql_file_error(error)
            return TransformResult(sql_file=sql_file, success=False, error=error)

        if not sql_path.exists():
            error = f"SQL file not found: {sql_path}"
            if self.reporter:
                self.reporter.print_sql_file_not_found(sql_file, str(sql_path))
            return TransformResult(sql_file=sql_file, success=False, error=error)
        try:
            sql_content = sql_path.read_text()
            statements = conn.extract_statements(sql_content)
            for i, statement in enumerate(statements, 1):
                try:
                    conn.execute(statement)
                except duckdb.Error as e:
                    msg = f"Statement {i} failed: {e}"
                    if self.reporter:
                        self.reporter.print_sql_file_error(msg)
                    return TransformResult(sql_file=sql_file, success=False, error=msg)
                if self.reporter:
                    self.reporter.print_sql_statement_executed(i)
            if self.reporter:
                self.reporter.print_sql_file_success()
            return TransformResult(sql_file=sql_file, success=True)
        except (OSError, duckdb.Error) as e:
            error = str(e)
            if self.reporter:
                self.reporter.print_sql_file_error(error)
            return TransformResult(sql_file=sql_file, success=False, error=error)
