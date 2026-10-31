from . import edit as edit
from . import linter as linter
from . import stats as stats

__version__: str

class SheetrsError(Exception):
    """A workbook could not be read, analyzed or written."""
