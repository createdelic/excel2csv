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


def should_include_in_filter(worksheet, rownum, filter):
    if not filter:
        return True
    filter_cell = worksheet.cell(rownum, int(filter[0]))
    if does_filter_key_match(filter_cell.value, filter[2]):
        return False
    if not does_filter_key_match(filter_cell.value, filter[1]):
        raise Exception('Invalid filter value in sheet_name=' + worksheet.name + ' row=' + str(rownum))
    return True
