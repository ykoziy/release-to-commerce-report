#! python3
# parse.py - walks through a folder tree finding excel files with macros,
# extracts macros into src folder, that is created in the current working directory.

import os;
import shutil;
from oletools.olevba import VBA_Parser;


def parse(file_path):
    vba_folder = 'src'
    vba_parser = VBA_Parser(file_path);
    for _, _, vba_filename, vba_code in vba_parser.extract_macros():
        if not os.path.isdir(vba_folder):
            os.makedirs(vba_folder)
        with open(os.path.join(vba_folder, vba_filename), 'w', encoding='utf-8') as file:
            file.write(vba_code);
    vba_parser.close();

if __name__ == "__main__":
    for folder, subfolders, files in os.walk('.'):
        for item in subfolders:
            if item == 'src':
                shutil.rmtree(os.path.join(folder, item))

        for item in files:
            if item.endswith('.xlsm'):
                parse(os.path.join(folder, item))