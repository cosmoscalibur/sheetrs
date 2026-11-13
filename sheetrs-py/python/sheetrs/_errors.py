"""The exceptions of the package.

``Error`` is what every failure of the tools raises, except the ones that are Python's own: a file
that cannot be read or written is an ``OSError`` (``FileNotFoundError``, ...) and a call with an
argument that is not allowed is a ``ValueError`` or ``TypeError``. The classes that are also a
``ValueError`` are the ones that say an argument is wrong: ``except ValueError`` catches them.
"""


class Error(Exception):
    """A workbook could not be read, analyzed, calculated or written."""


class FormatError(Error, ValueError):
    """The workbook is not an XLSX, XLSM or ODS workbook, or its format was not given."""


class UnsupportedDateSystemError(FormatError):
    """The workbook numbers its dates from a date that is not one of the three date systems of
    LibreOffice Calc.

    SheetRS supports 1899-12-30 (the default, and the 1900 system of an XLSX), 1900-01-01 (StarCalc 1.0)
    and 1904-01-01. LibreOffice can store another null date (its API, a crafted ODS) and SheetRS refuses it,
    naming the date, instead of numbering the dates from it. It is raised when the workbook is read.
    """


class InvalidRangeError(Error, ValueError):
    """A range, a sheet scope or a name that is not valid for the workbook."""


class UnavailableFunctionError(Error):
    """A formula uses a function the calculation engine does not have."""


class PendingOperationsError(Error):
    """A query was asked of a ``Book`` that has operations nothing has run yet.

    The query would be about a workbook that does not exist yet: ``save()`` first and query the
    ``Book`` it returns, or run the query before adding the operations when it does not depend on them.
    """


class RecalculationWarning(UserWarning):
    """The engine could not calculate the workbook, so it was written marked to be recalculated.

    ``Book.save`` still writes the workbook as the other operations leave it, with the mark that makes
    the application that opens it recalculate every formula, and says so with this warning: the
    recalculation that is needed will be done by the spreadsheet program that opens the file later. The
    message gives the reason (a function the engine does not have, a reference it does not read).
    """
