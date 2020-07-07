import excel2csvlib.helpers as helpers

import csv
import json

import xlrd
import os
import pathlib
import argparse


def excel_to_csv(
        excel_file,
        csv_file_base_path,
        columns_to_copy,
        replacements,
        delimiter,
        comment_start,
        worksheet_path_sep,
        should_trim,
        quotechar,
        output_paths,
        column_filter):

    workbook = xlrd.open_workbook(excel_file)
    for sheet_name in workbook.sheet_names():
        print('Reading sheet \"' + sheet_name + '\"')
        worksheet = workbook.sheet_by_name(sheet_name)

        if output_paths:
            csv_file_relative_path = output_paths[sheet_name]
        else:
            csv_file_relative_path = sheet_name.replace(worksheet_path_sep, os.path.sep) + '.csv'

        csv_file_full_path = os.path.join(
            csv_file_base_path,
            csv_file_relative_path)
        pathlib.Path(os.path.dirname(csv_file_full_path)).mkdir(parents=True, exist_ok=True)

        with open(csv_file_full_path, 'w', newline='', encoding='utf-8') as csvfile:
            writetocsv = csv.writer(
                csvfile,
                delimiter=delimiter,
                quoting=csv.QUOTE_ALL,
                quotechar=quotechar)
            for rownum in range(worksheet.nrows):
                if worksheet.row_len(rownum) == 0:
                    continue

                first_cell = worksheet.cell(rownum, 0).value
                if isinstance(first_cell, str) and first_cell.startswith(comment_start):
                    continue

                if not helpers.should_include_in_filter(worksheet, rownum, column_filter):
                    continue

                writetocsv.writerow(
                    [helpers.process_cell(worksheet.cell(rownum, x).value, replacements, should_trim) for x in columns_to_copy]
                )

            print('Saved sheet \"' + sheet_name + '\" to ' + csv_file_full_path)


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', action='store', dest='infile', help='the input file', required=True)
    parser.add_argument('-o', action='store', dest='outpath', help='the output base path', required=True)
    parser.add_argument('-c', dest='columns', metavar='N', type=int, nargs='+', help='the columns to extract', required=True)
    parser.add_argument('-r', dest='replacements', metavar='C', nargs=2, action='append', help='characters to replace', required=False)
    parser.add_argument('-p', dest='configfile', help='config file', required=False)
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
    parser.add_argument('--comment', dest='comment_start', default=helpers.DEFAULT_COMMENT_START, help='character to start a comment', required=False)
    parser.add_argument('--path-sep', dest='worksheet_path_sep', default=helpers.DEFAULT_WORKSHEET_NAME_PATH_SEP,
                        help='character to indicate a path separator in a worksheet name', required=False)

    parser.add_argument('--version', action='version', version='%(prog)s ' + helpers.PROGRAM_VERSION)

    parser.set_defaults(unify_quotemarks=False, should_trim=False, quotechar=helpers.DEFAULT_QUOTECHAR)

    args = parser.parse_args()

    replacements = dict(args.replacements) if args.replacements else {}

    if args.unify_quotemarks:
        replacements['“'] = '"'
        replacements['”'] = '"'
        replacements['‘'] = "'"
        replacements['’'] = "'"

    output_paths = None
    if args.configfile:
        with open(args.configfile) as f:
            output_paths = json.load(f)

    excel_to_csv(
        excel_file=args.infile,
        csv_file_base_path=args.outpath,
        columns_to_copy=args.columns,
        replacements=replacements,
        delimiter=args.delimiter,
        comment_start=args.comment_start,
        worksheet_path_sep=args.worksheet_path_sep,
        should_trim=args.trim,
        quotechar=args.quotechar,
        output_paths=output_paths,
        column_filter=args.column_filter,
    )


if __name__ == '__main__':
    main()
