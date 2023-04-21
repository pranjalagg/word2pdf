from pathlib import Path
import sys
import re
import os

def extractTags(document, tags):
    for line in document.paragraphs:
        for word in line.text.split():
            regex_lst = re.findall("«.*»", word)
            try:
                tags[regex_lst[0]] = tags.get(regex_lst[0], 0) + 1
            except:
                pass
    return tags


def saveAsPDF(word, filepath, pdf_path):
    document = word.Documents.Open(str(filepath))
    document.SaveAs(str(pdf_path), FileFormat=17) # wdFromatPDF = 17
    document.Close(0)

def saveAsDocx(word, filepath):
    print(f" Converting {filepath.stem} to .docx")
    doc_file = filepath.parent / filepath.stem
    document = word.Documents.Open(str(filepath))
    document.SaveAs(str(doc_file) + ".docx", FileFormat=16) # wdFormatDocumentDefault = 16
    document.Close(0)

def resolvePath(in_path, out_path=None):
    in_path = Path(in_path).resolve()
    if out_path:
        out_path = Path(out_path).resolve()
    
    paths = {}
    if in_path.is_file():
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
        print("Please check your path and try again. Remember to remove the '\\' at the end if it is a folder")
        sys.exit(0)

    return paths

# def resolvePath(in_path):
#     in_path = Path(in_path).resolve()

#     paths = {}
#     if in_path.is_file():
#         print('Getting info from file ...')
#         paths['bulk'] = False
#         paths['input'] = str(in_path)

#     elif in_path.is_dir():
#         print('Inside the directory ...')
#         paths['bulk'] = True
#         paths['input'] = str(in_path)

#     else:
#         print("Please check your path and try again. Remember to remove the '\\' at the end if it is a folder")
#         sys.exit(0)

#     return paths