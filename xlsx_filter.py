from pathlib import Path
import openpyxl as xl
from openpyxl.descriptors.base import String
import langid
from math import floor
import argparse

langid.set_languages(('ru', 'en'))

parser = argparse.ArgumentParser(description="Remove english text from xlsx files", formatter_class=argparse.RawTextHelpFormatter, epilog='Levels of verbosity:\n-v\t\tPrint filenames\n-vv\t\tPrint detected language pairs\n-vvv\t\tPrint detected single lang cells')
parser.add_argument("path", metavar="P", nargs='+', help='Path to file(s) or directory')
parser.add_argument('-v', '--verbose', action='count', default=0)
parser.add_argument('-o', '--output', metavar="O", help='Path to the output file')
parser.add_argument('-x', '--overwrite', action='store_const', const=True, default=False, help='Overwrite input files')
parser.add_argument('-s', '--min-char-single', metavar='MS', type=int, default=10, help='Minimum number of characters required to delete a single language english cell (default: 10')

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
                return [text[:mid_a], text[mid_a + 1:]]
            else:
                return [text[:mid_b], text[mid_b + 1:]]
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

                # Check for Russian/English or English/Russian syntax
                if cell.value is not None and str(cell.value)[:1] != '=' and type(cell.value) is not int:
                    text = split_pair(str(cell.value))

                    if len(text) == 1:
                        if args.verbose > 2:
                            print("Single lang: " + text[0])
                        
                        if len(text[0]) > args.min_char_single:
                            lang = langid.classify(text[0])

                            if lang[0] == 'en':
                                cell.value = None

                    elif len(text) == 2:
                        # Check for ru-en pair
                        lang_a = langid.classify(text[0])
                        lang_b = langid.classify(text[1])

                        if lang_a[0] == 'ru' and lang_b[0] == 'en':
                            cell.value = text[0]
                        elif lang_a[0] == 'en' and lang_b[0] == 'ru':
                            cell.value = text[1]

                        if args.verbose > 1:
                            print("Dual lang [" + lang_a[0] + "/" + lang_b[0] + "]: " + cell.value)
                    else:
                        print("ERROR")
    if args.overwrite is True:
        wb.save(path)
    else:
        wb.save(args.output)

    wb.save(args.output)

def main():
    if (len(args.path) == 1 and args.output is None) or len(args.path) > 1 or args.overwrite is True:
        proceed = input('This operation will overwrite the selected files! Proceed? [Y/N]')
        if proceed.upper() == 'Y':
            args.overwrite = True
            for path in args.path:
                if path[-5:] == '.xlsx':
                    filter_xlsx(path)
                else:
                    path_obj = Path(path)
                    for file in path_obj.glob("**/*.xlsx"):
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
