import csv
import xlrd
import os
import pathlib
import argparse

PROGRAM_VERSION='1.0'

DEFAULT_COMMENT_START = '#'
DEFAULT_DELIMITER = '|'
DEFAULT_WORKSHEET_NAME_PATH_SEP = '>'


def replace_all(text, replacements):
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text


def excel_to_csv(excel_file, csv_file_base_path, columns_to_copy, replacements, delimiter, comment_start, worksheet_path_sep):
    workbook = xlrd.open_workbook(excel_file)
    for sheet_name in workbook.sheet_names():
        print('Reading sheet \"' + sheet_name + '\"')
        worksheet = workbook.sheet_by_name(sheet_name)

        csv_file_full_path = os.path.join(
            csv_file_base_path,
            sheet_name.replace(worksheet_path_sep, os.path.sep) + '.csv')
        pathlib.Path(os.path.dirname(csv_file_full_path)).mkdir(parents=True, exist_ok=True)

        csvfile = open(csv_file_full_path, 'w', newline='')
        writetocsv = csv.writer(csvfile, delimiter=delimiter, quoting=csv.QUOTE_ALL)
        for rownum in range(worksheet.nrows):
            if worksheet.row_len(rownum) == 0:
                continue

            if worksheet.cell(rownum, 0).value.startswith(comment_start):
                continue

            writetocsv.writerow(
                [replace_all(worksheet.cell(rownum, x).value, replacements) for x in columns_to_copy]
            )
        csvfile.close()

        print('Saved sheet \"' + sheet_name + '\" to ' + csv_file_full_path)


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument('-i', action='store', dest='infile', help='the input file', required=True)
    parser.add_argument('-o', action='store', dest='outpath', help='the output base path', required=True)
    parser.add_argument('-c', dest='columns', metavar='N', type=int, nargs='+', help='the columns to extract', required=True)
    parser.add_argument("-r", dest='replacements', metavar='C', nargs=2, action='append', help='characters to replace', required=False)

    parser.add_argument('--delimiter', dest='delimiter', default=DEFAULT_DELIMITER, help='character to separate columns', required=False)
    parser.add_argument('--comment', dest='comment_start', default=DEFAULT_COMMENT_START, help='character to start a comment', required=False)
    parser.add_argument('--path-sep', dest='worksheet_path_sep', default=DEFAULT_WORKSHEET_NAME_PATH_SEP,
                        help='character to indicate a path separator in a worksheet name', required=False)

    parser.add_argument('--version', action='version', version='%(prog)s ' + PROGRAM_VERSION)

    args = parser.parse_args()

    excel_to_csv(
        excel_file=args.infile,
        csv_file_base_path=args.outpath,
        columns_to_copy=args.columns,
        replacements=dict(args.replacements),
        delimiter=args.delimiter,
        comment_start=args.comment_start,
        worksheet_path_sep=args.worksheet_path_sep,
    )


if __name__ == '__main__':
    main()
