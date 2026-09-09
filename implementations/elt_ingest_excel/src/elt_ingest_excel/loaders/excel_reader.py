# src/elt_ingest_excel/loader/excel_reader.py
import re
import pandas as pd
from pathlib import Path


class ExcelReader:
    def __init__(
        self,
        file_path,
        sheet_name=0,
        header_row: int = 1,
        data_row: int | None = None,
        column_headers: list[str] | None = None,
        dtype=str,
    ):
        self.file_path = Path(file_path).expanduser()
        self.sheet_name = sheet_name
        self.header_row = int(header_row)
        self.data_row = int(data_row) if data_row is not None else None
        self.column_headers = [str(c) for c in column_headers] if column_headers else None
        self.dtype = dtype
        self.df = None

    @staticmethod
    def clean_column_name(name: str) -> str:
        """Clean a column name for database compatibility.

        Args:
            name: Original column name.

        Returns:
            Cleaned column name:
            - Lowercased
            - Trimmed (leading/trailing spaces removed)
            - Non-alphanumeric characters replaced with underscore
            - Multiple consecutive underscores collapsed to single
            - Leading/trailing underscores removed
        """
        if not isinstance(name, str):
            name = str(name)
        # Lowercase and trim
        cleaned = name.lower().strip()
        # Replace non-alphanumeric with underscore
        cleaned = re.sub(r'[^a-z0-9]', '_', cleaned)
        # Collapse multiple underscores to single
        cleaned = re.sub(r'_+', '_', cleaned)
        # Remove leading/trailing underscores
        cleaned = cleaned.strip('_')
        return cleaned

    @staticmethod
    def _normalize_sheet_name(name: str) -> str:
        return re.sub(r"\s+", " ", str(name)).strip().casefold()

    def _load_with_xlwings(self) -> pd.DataFrame:
        import xlwings as xw

        app = xw.App(visible=False, add_book=False)
        book = None
        try:
            book = app.books.open(str(self.file_path))
            sheet = None

            if isinstance(self.sheet_name, int):
                sheet = book.sheets[self.sheet_name]
            elif isinstance(self.sheet_name, str):
                try:
                    sheet = book.sheets[self.sheet_name]
                except Exception:
                    normalized_target = self._normalize_sheet_name(self.sheet_name)
                    normalized_matches = [
                        s
                        for s in book.sheets
                        if self._normalize_sheet_name(s.name) == normalized_target
                    ]
                    if len(normalized_matches) == 1:
                        sheet = normalized_matches[0]
                    elif len(book.sheets) == 1:
                        sheet = book.sheets[0]
                    else:
                        raise ValueError(
                            f"Worksheet named '{self.sheet_name}' not found. Available worksheets: {[s.name for s in book.sheets]}"
                        )
            else:
                raise TypeError(f"Unsupported sheet_name type: {type(self.sheet_name)}")

            used = sheet.used_range
            values = used.value
            if values is None:
                return pd.DataFrame()
            if not isinstance(values, list):
                values = [[values]]
            elif values and not isinstance(values[0], list):
                values = [values]

            used_row = int(getattr(used, "row", 1) or 1)
            used_col = int(getattr(used, "column", 1) or 1)

            pad_rows = max(0, used_row - 1)
            pad_cols = max(0, used_col - 1)

            if pad_cols:
                values = [[None] * pad_cols + list(r) for r in values]

            if pad_rows:
                empty_row = [None] * (len(values[0]) if values else 0)
                values = [empty_row[:] for _ in range(pad_rows)] + values

            raw = pd.DataFrame(values)
            return self._finalize_raw(raw)
        finally:
            if book is not None:
                book.close()
            app.quit()

    @staticmethod
    def _make_unique(columns: list[str]) -> list[str]:
        seen: dict[str, int] = {}
        out: list[str] = []
        for c in columns:
            base = c if c else "column"
            n = seen.get(base, 0)
            if n == 0:
                out.append(base)
            else:
                out.append(f"{base}_{n}")
            seen[base] = n + 1
        return out

    def _finalize_raw(self, raw: pd.DataFrame) -> pd.DataFrame:
        if raw.empty:
            return raw

        header_row = int(self.header_row)
        data_row = int(self.data_row) if self.data_row is not None else None

        if header_row < 0:
            raise ValueError(f"header_row must be >= 0, got: {header_row}")
        if data_row is not None and data_row < 1:
            raise ValueError(f"data_row must be >= 1, got: {data_row}")

        if header_row == 0:
            if data_row is None:
                data_row = 1
            if self.column_headers:
                col_names = [self.clean_column_name(c) for c in self.column_headers]
                col_names = self._make_unique(col_names)
                raw = raw.iloc[:, : len(col_names)]
            else:
                col_names = [f"column_{i+1}" for i in range(raw.shape[1])]
        else:
            if header_row > len(raw):
                raise ValueError(
                    f"header_row={header_row} is out of bounds for worksheet '{self.sheet_name}'"
                )
            if data_row is None:
                data_row = header_row + 1
            if data_row <= header_row:
                raise ValueError(
                    f"data_row={data_row} must be greater than header_row={header_row} for worksheet '{self.sheet_name}'"
                )

            header_values = raw.iloc[header_row - 1].tolist()
            col_names = [self.clean_column_name(v) for v in header_values]
            col_names = self._make_unique(col_names)

        if data_row is None:
            data_row = 1
        df = raw.iloc[data_row - 1 :].copy()
        df.columns = col_names[: df.shape[1]]
        df = df.where(df.notna(), "")
        df = df.astype(str)
        df.dropna(how="all", axis=0, inplace=True)
        return df

    def load(self):
        if not self.file_path.exists():
            raise FileNotFoundError(f"File not found: {self.file_path}")
        try:
            raw = pd.read_excel(
                self.file_path,
                sheet_name=self.sheet_name,
                header=None,
                dtype=self.dtype,
                na_filter=False,
            )
        except ValueError as e:
            msg = str(e)
            if (
                "Worksheet named" in msg
                or "Worksheet index" in msg
                or "0 worksheets found" in msg
            ):
                if isinstance(self.sheet_name, str):
                    try:
                        import openpyxl

                        wb = openpyxl.load_workbook(
                            self.file_path, read_only=True, data_only=True
                        )
                        available = list(wb.sheetnames)
                        wb.close()
                        if available and self.sheet_name not in available:
                            normalized_target = self._normalize_sheet_name(self.sheet_name)
                            normalized_matches = [
                                n
                                for n in available
                                if self._normalize_sheet_name(n) == normalized_target
                            ]
                            if len(normalized_matches) == 1:
                                raw = pd.read_excel(
                                    self.file_path,
                                    sheet_name=normalized_matches[0],
                                    header=None,
                                    dtype=self.dtype,
                                    na_filter=False,
                                )
                            else:
                                raw = self._load_with_xlwings()
                        else:
                            raw = self._load_with_xlwings()
                    except Exception:
                        raw = self._load_with_xlwings()
                else:
                    raw = self._load_with_xlwings()
            else:
                raise

        self.df = self._finalize_raw(raw)
        if "department" in self.df.columns and "department_1" not in self.df.columns:
            self.df["department_1"] = self.df["department"]
        return self.df

    def preview(self, n=5, full=True):
        if self.df is None:
            print("No data loaded.")
            return
        if full:
            # Show all columns without truncation
            with pd.option_context(
                'display.max_columns', None,
                'display.width', 10000,
                'display.max_colwidth', 100
            ):
                print(self.df.head(n))
        else:
            print(self.df.head(n))
