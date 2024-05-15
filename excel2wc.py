import excel2csvlib.helpers as helpers

import openpyxl
import argparse
import string


def replace_control_codes_with_whitespace(s, open_symbol, close_symbol):
    while True:
        open_bracket_pos = s.find(open_symbol)
        if open_bracket_pos < 0:
            return s
        close_bracket_pos = s.find(close_symbol)
        s = s[0:open_bracket_pos] + ' ' + s[close_bracket_pos+1:]


def get_words(s):
    s = s.strip()
    s = replace_control_codes_with_whitespace(s, '[', ']').strip()
    s = replace_control_codes_with_whitespace(s, '{', '}').strip()
    s = s.translate(str.maketrans('', '', string.punctuation))  # remove punctuation
    s = list(filter(None, s.split(' ')))
    return s

def excel_to_wc(
        excel_file,
        columns_to_count,
        replacements,
        comment_start,
        column_filter,
        ignore_if_columns_match,
        ignore_if_column_empty,
        ignore_if_column_not_empty,
):

    total_word_count = 0
    workbook = openpyxl.load_workbook(excel_file, data_only=True)
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        sheet_word_count = 0

        for row in worksheet:
            if helpers.is_row_empty(row, columns_to_count):
                continue

            if helpers.is_row_ignored(row, comment_start):
                continue

            if ignore_if_column_empty and helpers.is_cell_empty(row[ignore_if_column_empty]):
                continue

            if ignore_if_column_not_empty and helpers.is_cell_not_empty(row[ignore_if_column_not_empty]):
                continue

            if ignore_if_columns_match:
                should_ignore = False
                first_match = True
                match_contents = None
                for column_to_match in ignore_if_columns_match:
                    cell_contents = helpers.process_cell(row[column_to_match], replacements, True)
                    if (not first_match) and (cell_contents == match_contents):
                        should_ignore = True
                        break
                    first_match = False
                    match_contents = cell_contents
                if should_ignore:
                    continue

            if not helpers.should_include_in_filter(row, column_filter):
                continue

            for column_to_count in columns_to_count:
                cell_contents = helpers.process_cell(row[column_to_count], replacements, True)
                sheet_word_count += len(get_words(cell_contents))

        print(sheet_name + '\t' + str(sheet_word_count))
        total_word_count += sheet_word_count

    print('----------')
    print('TOTAL\t' + str(total_word_count))


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', action='store', dest='infile', help='the input file', required=True)
    parser.add_argument('-c', dest='columns', metavar='N', type=int, nargs='+', help='the columns to count words in', required=True)
    parser.add_argument('-r', dest='replacements', metavar='C', nargs=2, action='append', help='characters to replace', required=False)
    parser.add_argument('-f', dest='column_filter', nargs=3, help='filter only rows containing a value', required=False)
    parser.add_argument('--ignore-if-columns-match',
                        dest='ignore_if_columns_match',
                        metavar='N',
                        type=int,
                        nargs='+',
                        help='ignore if columns match',
                        required=False)
    parser.add_argument('--ignore-if-column-empty', type=int, dest='ignore_if_column_empty', help='ignore if column is empty', required=False)
    parser.add_argument('--ignore-if-column-not-empty', type=int, dest='ignore_if_column_not_empty', help='ignore if column is not empty', required=False)

    parser.add_argument('--comment', dest='comment_start', default=helpers.DEFAULT_COMMENT_START, help='character to start a comment', required=False)

    parser.add_argument('--version', action='version', version='%(prog)s ' + helpers.PROGRAM_VERSION)

    args = parser.parse_args()

    replacements = dict(args.replacements) if args.replacements else {}

    excel_to_wc(
        excel_file=args.infile,
        columns_to_count=args.columns,
        replacements=replacements,
        comment_start=args.comment_start,
        column_filter=args.column_filter,
        ignore_if_columns_match=args.ignore_if_columns_match,
        ignore_if_column_empty=args.ignore_if_column_empty,
        ignore_if_column_not_empty=args.ignore_if_column_not_empty,
    )


if __name__ == '__main__':
    main()
