import openpyxl.cell.cell
from openpyxl.cell.read_only import EmptyCell

PROGRAM_VERSION = '2023.8.8'

DEFAULT_COMMENT_START = '#'
DEFAULT_DELIMITER = '|'
DEFAULT_WORKSHEET_NAME_PATH_SEP = '>'
DEFAULT_QUOTECHAR = '@'


def is_row_empty(row, column_indexes):
    if all(is_cell_empty(cell) for cell in row):
        return True

    if all(is_cell_empty(row[column_index]) for column_index in column_indexes):
        return True

    return False


def is_cell_empty(cell):
    return isinstance(cell, EmptyCell) \
           or (cell.value is None) \
           or (cell.data_type == 's' and (not cell.value.strip()))


def is_cell_not_empty(cell):
    return (not isinstance(cell, EmptyCell)) \
           and (cell.value is not None) \
           and (cell.data_type == 's' and cell.value.strip())


def is_row_ignored(row, comment_start):
    cell = row[0]
    return (cell.data_type == 's') and cell.value.startswith(comment_start)


def process_cell(cell, replacements, should_trim):
    cell_value = cell.value
    if cell.data_type == 's':
        for k, v in replacements.items():
            cell_value = cell_value.replace(k, v)
        if should_trim:
            cell_value = cell_value.strip()
    return cell_value


def does_filter_key_match(a, b):
    a = a.lower().strip()
    b = b.lower().strip()
    return a == b


def should_include_in_filter(row, column_filter):
    if not column_filter:
        return True
    column_index = int(column_filter[0])
    filter_cell_value = row[column_index].value
    if (not filter_cell_value) or (not filter_cell_value.strip()) or does_filter_key_match(filter_cell_value, column_filter[2]):
        return False
    if not does_filter_key_match(filter_cell_value, column_filter[1]):
        raise Exception('Invalid filter value: ' + filter_cell_value)
    return True
