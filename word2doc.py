import argparse
from pathlib import Path
import os
import win32com.client
import sys
from tqdm import tqdm


def convertToPdf(paths):
    word_instance = win32com.client.Dispatch("Word.Application")
    wdFromatPDF = 17

    if paths['bulk']:
        print("Converting files from the given folder")
        for filepath in tqdm(Path(paths['input']).glob('*.doc*')):
            pdf_path = Path(paths['output']) / (str(filepath.stem) + ".pdf")
            document = word_instance.Documents.Open(str(filepath))
            document.SaveAs(str(pdf_path), FileFormat=wdFromatPDF)
            document.Close(0)
    
    else:
        print("Converting ...")
        filepath = Path(paths['input'])
        pdf_path = Path(paths['output'])
        document = word_instance.Documents.Open(str(filepath))
        document.SaveAs(str(pdf_path), FileFormat=wdFromatPDF)
        document.Close(0)

def identify_path(in_path, out_path=None):
    # Resolve paths to handle relative paths
    in_path = Path(in_path).resolve()
    if out_path is not None:
        out_path = Path(out_path).resolve()

    # print("in_path ", in_path)
    # print("Folder ", in_path.is_dir()) 
    # print("out_path ", out_path)
    # print("Folder ", out_path.is_dir())

    paths = {}
    # Condition when the path is a file
    if os.path.isfile(in_path):
        print("Identified input as a file")
        paths['bulk'] = False
        paths['input'] = str(in_path)

        if out_path and os.path.isdir(out_path):
            out_path = os.path.join(out_path, in_path.stem) + ".pdf"
        elif out_path:
            print("Output path does not exist")
            sys.exit(0)
        else:
            out_path = os.path.join(in_path.parent, in_path.stem) + ".pdf"

        paths['output'] = out_path

    # Condition when the path is a folder
    elif os.path.isdir(in_path):
        print("Identified input as a folder")
        paths['bulk'] = True
        paths['input'] = str(in_path)

        if out_path and os.path.isdir(out_path):
            pass
        elif out_path:
            print("Path does not exist")
            sys.exit(0)
        else:
            out_path = str(in_path)

        paths['output'] = out_path
    
    else:
        print("Please check your path and try again")
        sys.exit(0)

    return paths

def main():

    # Initialise parser
    parser = argparse.ArgumentParser(description="Program to convert word files to pdf via CLI")

    # Add arguments
    parser.add_argument('inpath', help="Path of the file or folder to be converted.")
    parser.add_argument('outpath', nargs='?', help="Path of the folder to store the converted files. Defaulted to inpath.")

    args = parser.parse_args()

    # print(args.outpath)
    paths = identify_path(args.inpath, args.outpath)

    convertToPdf(paths)

main()