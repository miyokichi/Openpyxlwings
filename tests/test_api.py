from openpyxlwings import ExcelReader, ExcelWorkbook, ExcelWriter


def test_legacy_class_names_point_to_unified_class() -> None:
    assert ExcelReader is ExcelWorkbook
    assert ExcelWriter is ExcelWorkbook
