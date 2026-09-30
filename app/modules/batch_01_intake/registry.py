from .find_duplicates import run as find_duplicates
from .import_patients import run as import_patients
from .inspect_workbook import run as inspect_workbook
from .system_check import run as system_check
from .validate_patients import run as validate_patients

HANDLERS = {
    "b1_system_check": system_check,
    "b1_inspect_workbook": inspect_workbook,
    "b1_import_patients": import_patients,
    "b1_validate_patients": validate_patients,
    "b1_find_duplicates": find_duplicates,
}
