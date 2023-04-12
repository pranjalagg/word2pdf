import argparse
import docx
import re
import os
from pathlib import Path
import win32com.client
import sys

def getTags(paths, rm):
    word = win32com.client.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16

    if paths['bulk']:
        for filepath in sorted(Path(paths['input']).glob("*.doc*")):
            if str(filepath).endswith(".doc"):
                doc_file = filepath.parent / filepath.stem
                # print(doc_file, ".docx")
                document = word.Documents.Open(str(filepath))
                document.SaveAs(str(doc_file) + ".docx", FileFormat=wdFormatDocumentDefault)
                document.Close(0)
                if rm:
                    os.remove(str(doc_file) + ".doc")
        
        for filepath in sorted(Path(paths['input']).glob("*.docx")):
            # print(filepath)
            document = docx.Document(str(filepath))


def resolvePath(in_path):
    in_path = Path(in_path).resolve()

    paths = {}
    if in_path.is_file():
        print('Getting info from file ...')
        paths['bullk'] = False
        paths['input'] = str(in_path)

    elif in_path.is_dir():
        print('Inside the directory ...')
        paths['bulk'] = True
        paths['input'] = str(in_path)

    else:
        print("Please check your path and try again")
        sys.exit(0)

    return paths

def main():
    # Initialise parser
    parser = argparse.ArgumentParser(description="Tool to extract specific tags from word document")

    # Add arguments
    parser.add_argument('inpath', help="Path of the file or folder.")
    parser.add_argument('--rm', default=False, help="Set to True, if in-place operation is needed")

    args = parser.parse_args()

    # print(args.rm)
    paths = resolvePath(args.inpath)
    getTags(paths, args.rm)

main()