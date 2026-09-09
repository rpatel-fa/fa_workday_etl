"""VBA Macro utilities for Excel workbooks."""

__all__ = ["run_validation", "list_vba_entry_points"]


def __getattr__(name: str):
    """Lazy loading of module exports."""
    if name == "run_validation":
        from elt_ingest_excel.macro.vba_runner import run_validation
        return run_validation
    if name == "list_vba_entry_points":
        from elt_ingest_excel.macro.vba_analyzer import list_vba_entry_points
        return list_vba_entry_points
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
