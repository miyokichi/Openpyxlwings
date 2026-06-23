from openpyxlwings import (
    ExcelFormat,
    ExcelReader,
    ExcelWorkbook,
    ExcelWriter,
    ExtractedMatch,
    Placeholder,
    TablePattern,
)


def test_legacy_class_names_point_to_unified_class() -> None:
    assert ExcelReader is ExcelWorkbook
    assert ExcelWriter is ExcelWorkbook


def test_format_api_is_public() -> None:
    assert ExcelFormat.__name__ == "ExcelFormat"
    assert ExtractedMatch.__name__ == "ExtractedMatch"
    assert Placeholder.__name__ == "Placeholder"
    assert TablePattern.__name__ == "TablePattern"
