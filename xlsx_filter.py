from pathlib import Path
import openpyxl as xl
from openpyxl.descriptors.base import String
import langid
from math import floor
import argparse
import re

langid.set_languages(('ru', 'en'))

parser = argparse.ArgumentParser(description="Remove english text from xlsx files", formatter_class=argparse.RawTextHelpFormatter, epilog='Levels of verbosity:\n-v\t\tPrint filenames\n-vv\t\tPrint detected language pairs\n-vvv\t\tPrint detected single lang cells')
parser.add_argument("path", metavar="P", nargs='+', help='Path to file(s) or directory')
parser.add_argument('-v', '--verbose', action='count', default=0)
parser.add_argument('-o', '--output', metavar="O", help='Path to the output file')
parser.add_argument('-x', '--overwrite', action='store_const', const=True, default=False, help='Overwrite input files')
regex_group = parser.add_mutually_exclusive_group()
regex_group.add_argument('-r', '--regex', metavar='R', help='RegEx expression to always delete')
regex_group.add_argument('-R', '--regex-v', metavar='RV', help='Same as -r, but with verbose output')
parser.add_argument('-s', '--min-single', metavar='S', type=int, help='Minimum number of characters required to delete a single language english cell')

args = parser.parse_args()


# Very dumb function to find the middle '/' in a string
def split_pair(text: String):
    slashes = text.count("/")

    if slashes == 0:
        return [text]
    elif slashes == 1:
        return text.split("/")
    else:
        # Count "/" chars
        indexes = []
        i = 0
        while i < len(text):
            i = text.find("/", i)
            if i == -1:
                break
            indexes.append(i)
            i += 1
        
        # Find middle
        if (len(indexes) % 2) == 0:
            mid_a = indexes[floor(len(indexes) / 2)]
            mid_b = indexes[floor(len(indexes) / 2) - 1]

            if len(text) / 2 - mid_a > len(text) / 2 - mid_b:
                return [text[:mid_b], text[mid_b + 1:]]
            else:
                return [text[:mid_a], text[mid_a + 1:]]
        else:
            middle_index = indexes[floor(len(indexes) / 2)]
            return [text[0:middle_index], text[middle_index + 1:]]

def filter_xlsx(path):
    # Load workbooks
    wb = xl.load_workbook(path)

    if args.verbose > 0:
        print(path)

    # Iterate sheets
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:

                # Delete RegEx matches
                if args.regex_v is not None:
                    args.regex = args.regex_v
                if args.regex is not None:
                    match = re.match(args.regex, str(cell.value))
                    if match:
                        if args.regex_v is not None:
                            print('RegEx matched: ' + str(cell.value))
                        cell.value = ''
                        continue

                if cell.value is not None and str(cell.value)[:1] != '=' and type(cell.value) is not int:
                    text = split_pair(str(cell.value))

                    # Single language cell
                    if len(text) == 1 and args.min_single is not None:
                        if args.verbose > 2:
                            print("Single lang: " + text[0])
                        
                        if len(text[0]) > args.min_single:
                            lang = langid.classify(text[0])

                            if lang[0] == 'en':
                                cell.value = None

                    # Check for ru-en pair
                    elif len(text) == 2:
                        lang_a = langid.classify(text[0])
                        lang_b = langid.classify(text[1])

                        if args.verbose > 1:
                            print("Dual lang [" + lang_a[0] + "/" + lang_b[0] + "]: " + cell.value)

                        if lang_a[0] == 'ru' and lang_b[0] == 'en':
                            cell.value = text[0]
                        elif lang_a[0] == 'en' and lang_b[0] == 'ru':
                            cell.value = text[1]
                        elif lang_a[0] == 'en' and lang_b[0] == 'en' and args.min_single is not None:
                            if len(text[0]) > args.min_single:
                                cell.value = None

    if args.overwrite is True:
        wb.save(path)
    else:
        wb.save(args.output)

def main():
    if (len(args.path) == 1 and args.output is None) or len(args.path) > 1 or args.overwrite is True:
        proceed = input('This operation will overwrite the selected files! Proceed? [Y/N]')
        if proceed.upper() == 'Y':
            args.overwrite = True
            for path in args.path:
                if path[-5:] == '.xlsx' or path[-5:] == '.xlsm' or path[-5:] == '.xls':
                    filter_xlsx(path)
                else:
                    path_obj = Path(path)
                    for file in path_obj.glob('**/*'):
                        if file.name.endswith(('.xlsx', '.xlsm', '.xls')):
                            filter_xlsx(file)
        else:
            print('Stopping...')
    elif len(args.path) == 1 and args.output is not None:
        args.overwrite = False
        filter_xlsx(args.path[0])
    else:
        print('Invalid combination of arguments!')


if __name__ == '__main__':
    main()
