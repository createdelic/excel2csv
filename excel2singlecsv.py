import excel2csvlib.helpers as helpers

import csv

import openpyxl
import argparse


def excel_to_csv(
        excel_file,
        out_csv_file,
        columns_to_copy,
        replacements,
        delimiter,
        comment_start,
        should_trim,
        quotechar,
        column_filter,
        ignore_if_column_empty):

    workbook = openpyxl.load_workbook(excel_file, read_only=True)
    with open(out_csv_file, 'w', newline='', encoding='utf-8') as csvfile:
        writetocsv = csv.writer(
            csvfile,
            delimiter=delimiter,
            quoting=csv.QUOTE_ALL,
            quotechar=quotechar)

        for sheet_name in workbook.sheetnames:
            print('Reading sheet \"' + sheet_name + '\"')
            worksheet = workbook[sheet_name]

            for row in worksheet:
                if helpers.is_row_empty(row, columns_to_copy):
                    continue

                if helpers.is_row_ignored(row, comment_start):
                    continue

                if ignore_if_column_empty and helpers.is_cell_empty(row[ignore_if_column_empty]):
                    continue

                if not helpers.should_include_in_filter(row, column_filter):
                    continue

                data = [helpers.process_cell(row[x], replacements, should_trim) for x in
                        columns_to_copy]
                data.insert(0, sheet_name)

                writetocsv.writerow(data)

    print('Saved to: ' + out_csv_file)


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', action='store', dest='infile', help='the input file', required=True)
    parser.add_argument('-o', action='store', dest='outfile', help='the output file', required=True)
    parser.add_argument('-c', dest='columns', metavar='N', type=int, nargs='+', help='the columns to extract', required=True)
    parser.add_argument('-r', dest='replacements', metavar='C', nargs=2, help='characters to replace', required=False)
    parser.add_argument('-f', dest='column_filter', nargs=3, help='filter only rows containing a value', required=False)

    parser.add_argument("--unify-quotemarks",
                        dest='unify_quotemarks',
                        help='should left/right quotation marks be replaced with standard quote',
                        action='store_true',
                        required=False)
    parser.add_argument("--trim",
                        dest='trim',
                        help='should trim the contents of each cell',
                        action='store_true',
                        required=False)

    parser.add_argument('--quotechar', action='store', dest='quotechar', help='the quote character', required=False)
    parser.add_argument('--delimiter', dest='delimiter', default=helpers.DEFAULT_DELIMITER, help='character to separate columns', required=False)
    parser.add_argument('--ignore-if-column-empty', type=int, dest='ignore_if_column_empty',
                        help='ignore if column is empty', required=False)
    parser.add_argument('--comment', dest='comment_start', default=helpers.DEFAULT_COMMENT_START, help='character to start a comment', required=False)

    parser.add_argument('--version', action='version', version='%(prog)s ' + helpers.PROGRAM_VERSION)

    parser.set_defaults(unify_quotemarks=False, should_trim=False, quotechar=helpers.DEFAULT_QUOTECHAR)

    args = parser.parse_args()

    replacements = dict(args.replacements) if args.replacements else {}

    if args.unify_quotemarks:
        replacements['“'] = '"'
        replacements['”'] = '"'
        replacements['‘'] = "'"
        replacements['’'] = "'"

    excel_to_csv(
        excel_file=args.infile,
        out_csv_file=args.outfile,
        columns_to_copy=args.columns,
        replacements=replacements,
        delimiter=args.delimiter,
        comment_start=args.comment_start,
        should_trim=args.trim,
        quotechar=args.quotechar,
        column_filter=args.column_filter,
        ignore_if_column_empty=args.ignore_if_column_empty,
    )


if __name__ == '__main__':
    main()
