from pathlib import Path


class OrderReader:
    def __init__(self, base_path: Path):
        self.base_path = Path(base_path)
        self.order_file = self.base_path / "order.txt"

    def exists(self) -> bool:
        return self.order_file.exists()

    def read(self) -> list[str]:
        content = self.order_file.read_text()
        files: list[str] = []
        for line in content.strip().split("\n"):
            s = line.strip()
            if s and not s.startswith("#"):
                files.append(s)
        return files
