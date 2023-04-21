import argparse
from pathlib import Path
import win32com.client
from tqdm import tqdm
import pyperclip as pc
import helpers as hp


def convertToPdf(paths):
    word = win32com.client.Dispatch("Word.Application")
    wdFromatPDF = 17

    str_to_copy = ""

    if paths['bulk']:
        print("Converting files from the given folder")
        for filepath in tqdm(sorted(Path(paths['input']).glob('*.doc*'))):
            pdf_path = Path(paths['output']) / (str(filepath.stem) + ".pdf")
            hp.saveAsPDF(word, filepath, pdf_path)
            str_to_copy += str(pdf_path) + "\n"
            pc.copy(str_to_copy)
    
    else:
        print("Converting ...")
        filepath = Path(paths['input'])
        pdf_path = Path(paths['output'])
        hp.saveAsPDF(word, filepath, pdf_path)
        str_to_copy += str(pdf_path)
        pc.copy(str_to_copy)

def main():

    # Initialise parser
    parser = argparse.ArgumentParser(description="Program to convert word files to pdf via CLI")

    # Add arguments
    parser.add_argument('inpath', help="Path of the file or folder to be converted.")
    parser.add_argument('outpath', nargs='?', help="Path of the folder to store the converted files. Defaulted to inpath.")

    args = parser.parse_args()

    # print(args.outpath)
    paths = hp.resolvePath(args.inpath, args.outpath)

    convertToPdf(paths)

main()