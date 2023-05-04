import argparse
import docx
import re
import os
from pathlib import Path
import win32com.client
from tqdm import tqdm
import helpers as hp

def storeInfo(f, tags):
    for tag in list(tags.keys()):
        temp_tag = re.findall("«([^»]*)", tag)
        f.write(temp_tag[0] + "\n")

def getTags(paths, rm):
    word = win32com.client.Dispatch("Word.Application")
    # wdFormatDocumentDefault = 16

    if paths['bulk']:
        hp.convertDocFilesToDocx(paths['input'])
        # for filepath in tqdm(sorted(Path(paths['input']).glob("*.doc"))):
            # hp.saveAsDocx(word, filepath)
            # doc_file = filepath.parent / filepath.stem
            # if rm:
            #     os.remove(str(doc_file) + ".doc")
        
        f = open('Info.txt', "w+")
        for filepath in tqdm(sorted(Path(paths['input']).glob("*.docx"))):
            f.write("\n---- " + str(filepath.stem) + " ----\n")
            document = docx.Document(str(filepath))

            tags = {}
            tags = hp.extractTags(document, tags)
            storeInfo(f, tags)
        f.close()

    else:
        filepath = Path(paths['input'])
        if str(filepath).endswith(".doc") or str(filepath).endswith(".DOC"):
            # print("Here")
            hp.saveAsDocx(word, filepath)
            if rm:
                os.remove(str(filepath.parent / filepath.stem) + ".doc")
        
        f = open('Info.txt', "w+")
        f.write("\n---- " + str(filepath.stem) + " ----\n")
        document = docx.Document(str(filepath.parent / filepath.stem) + ".docx")

        tags = {}
        tags = hp.extractTags(document, tags)
        storeInfo(f, tags)
        f.close()

def main():
    # Initialise parser
    parser = argparse.ArgumentParser(description="Tool to extract specific tags from word document")

    # Add arguments
    parser.add_argument('inpath', help="Path of the file or folder.")
    parser.add_argument('--rm', default=False, help="Set to True, if in-place operation is needed")

    args = parser.parse_args()

    # print(args.rm)
    paths = hp.resolvePath(args.inpath, None)
    getTags(paths, args.rm)

main()