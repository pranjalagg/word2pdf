from pathlib import Path
import sys

def resolvePath(in_path):
    in_path = Path(in_path).resolve()

    paths = {}
    if in_path.is_file():
        print('Getting info from file ...')
        paths['bulk'] = False
        paths['input'] = str(in_path)

    elif in_path.is_dir():
        print('Inside the directory ...')
        paths['bulk'] = True
        paths['input'] = str(in_path)

    else:
        print("Please check your path and try again. Remember to remove the '\\' at the end if it is a folder")
        sys.exit(0)

    return paths