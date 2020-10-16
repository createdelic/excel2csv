PROGRAM_VERSION='1.0'

DEFAULT_COMMENT_START = '#'
DEFAULT_DELIMITER = '|'
DEFAULT_WORKSHEET_NAME_PATH_SEP = '>'
DEFAULT_QUOTECHAR = '@'


def process_cell(text, replacements, should_trim):
    for k, v in replacements.items():
        text = text.replace(k, v)
    if should_trim:
        text = text.strip()
    return text


def does_filter_key_match(a, b):
    a = a.lower().strip()
    b = b.lower().strip()
    return a == b


def should_include_in_filter(worksheet, rownum, column_filter):
    if not column_filter:
        return True
    filter_cell = worksheet.cell(rownum, int(column_filter[0]))
    if does_filter_key_match(filter_cell.value, column_filter[2]):
        return False
    if not does_filter_key_match(filter_cell.value, column_filter[1]):
        raise Exception('Invalid filter value in sheet_name=' + worksheet.name
                        + ' row=' + str(rownum)
                        + ' value=' + filter_cell.value)
    return True


def is_comment_row(worksheet, rownum, comment_start):
    first_cell = worksheet.cell(rownum, 0).value
    return isinstance(first_cell, str) and first_cell.startswith(comment_start)