"""Package-specific exceptions."""


class OpenpyxlWingsError(Exception):
    """Base exception for openpyxlwings."""


class SheetNotFoundError(OpenpyxlWingsError):
    """Raised when a requested worksheet does not exist."""


class ExcelWriteError(OpenpyxlWingsError):
    """Raised when Excel/xlwings cannot complete a write operation."""


class BorderTableError(OpenpyxlWingsError):
    """Base exception for bordered-table operations."""


class BorderTableNotFoundError(BorderTableError):
    """Raised when a bordered table cannot be found from a start cell."""


class BorderTableShapeError(BorderTableError):
    """Raised when a bordered table is malformed or edited inconsistently."""


class FormatError(OpenpyxlWingsError):
    """Base exception for Excel format definitions and extraction."""


class FormatDefinitionError(FormatError):
    """Raised when an Excel format definition is invalid."""


class FormatMatchError(FormatError):
    """Raised when a candidate table cannot be matched safely."""


class FormatValueError(FormatMatchError):
    """Raised when an extracted value violates a placeholder definition."""
