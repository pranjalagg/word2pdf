import argparse
from pathlib import Path
import os

def identify_path(in_path, out_path=None):
    # Resolve paths to handle relative paths
    in_path = Path(in_path).resolve()
    if out_path is not None:
        out_path = Path(out_path).resolve()

    dic = {}
    # Condition when the path is a file
    if os.path.isfile(in_path):
        print("Path is a file.")
        dic['bulk'] = False
        dic['input'] = str(in_path)

        if out_path and os.path.isdir(out_path):
            out_path = os.path.join(out_path, in_path.stem) + ".pdf"
        elif out_path:
            print("Path does not exist")
            sys.exit(0)
        else:
            out_path = os.path.join(in_path.parent, in_path.stem) + ".pdf"

        dic['output'] = out_path

    # Condition when the path is a folder
    else:
        print("Path is a folder.")
        dic['bulk'] = True
        dic['input'] = in_path

def main():

    # Initialise parser
    parser = argparse.ArgumentParser(description="Program to convert word files to pdf via CLI")

    # Add arguments
    parser.add_argument('inpath', help="Path of the file or folder to be converted.")
    parser.add_argument('outpath', nargs='?', help="Path of the folder to store the converted files. Defaulted to inpath.")

    args = parser.parse_args()

    identify_path(args.inpath, args.outpath)

main()