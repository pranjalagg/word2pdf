import argparse
import docx
import re
import os
from pathlib import Path
import win32com.client
import sys
from tqdm import tqdm

def replaceTags(document, filepath):
    for i in range(len(document.paragraphs)):
        line_str = document.paragraphs[i].text
        if line_str.endswith('<br>'):
            continue
        line_str += '<br>'
        document.paragraphs[i].text = line_str
        document.save(str(filepath))

def storeInfo(f, tags):
    for tag in list(tags.keys()):
        # temp_tag = re.findall("«([^»]*)", tag)
        # f.write(temp_tag[0] + "\n")
        f.write(tag + "\n")

def getTags(paths, rm):
    word = win32com.client.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16

    if paths['bulk']:
        for filepath in tqdm(sorted(Path(paths['input']).glob("*.doc"))):
            # if str(filepath).endswith(".doc"):
            doc_file = filepath.parent / filepath.stem
            # print(doc_file, ".docx")
            document = word.Documents.Open(str(filepath))
            document.SaveAs(str(doc_file) + ".docx", FileFormat=wdFormatDocumentDefault)
            document.Close(0)
            if rm:
                os.remove(str(doc_file) + ".doc")
        
        f = open('Info.txt', "w+")
        for filepath in tqdm(sorted(Path(paths['input']).glob("*.docx"))):
            # print(filepath)
            f.write("\n---- " + str(filepath.stem) + " ----\n")
            document = docx.Document(str(filepath))

            tags = {}
            for i in range(len(document.paragraphs)):
                replaceTags(document, filepath)
                # line_str = document.paragraphs[i].text
                # if line_str.endswith('<br>'):
                #     continue
                # line_str += '<br>'
                # document.paragraphs[i].text = line_str
                # document.save(str(filepath))
                # print(line.text)
                # word_lst = []
                # print(type(line.text))
                # print(line.text)
                # line_str = line.text
                # line_str = line_str.replace('^p', '<br>')
                # line.text = line_str
                # f.write(line_str)
                # break
                # for word in line.text.split():
                #     # word_lst.append(word)
                #     # word_lst.extend(re.findall("«.*»", word))

                #     temp = re.findall("«.*»", word)
                #     # print(word, temp)
                #     try:
                #         tags[temp[0]] = tags.get(temp[0], 0) + 1
                #     except:
                #         # print('---ERROR---')
                #         pass
                
                # print(word_lst)
            # print(tags)
            # storeInfo(f, tags)
            # if rm:
            #     os.remove(str(filepath))
            # break
        f.close()

    else:
        filepath = Path(paths['input'])
        # print(str(filepath))
        if str(filepath).endswith(".doc"):
            # print("Here")
            document = word.Documents.Open(str(filepath))
            document.SaveAs(str(filepath.parent / filepath.stem) + ".docx", FileFormat=wdFormatDocumentDefault)
            document.Close(0)
            if rm:
                os.remove(str(filepath.parent / filepath.stem) + ".doc")
        
        f = open('Info.txt', "w+")
        f.write("\n---- " + str(filepath.stem) + " ----\n")
        document = docx.Document(str(filepath.parent / filepath.stem) + ".docx")

        # tags = {}
        for i in range(len(document.paragraphs)):
            replaceTags(document, filepath)
            # line_str = document.paragraphs[i].text
            # words_lst.append(line_str + '---//--')
            # line_str = line_str.replace('\\n', '<br>')
            # line_str += '<br>'
            # document.paragraphs[i].text = line_str
            # document.save(str(filepath))
            # print(line_str)

            # for word  in line.text.split():
            #     temp = re.findall("«.*»", word)
            #     try:
            #         tags[temp[0]] = tags.get(temp[0], 0) + 1
            #     except:
            #         pass
        
        # storeInfo(f, tags)
        # if rm:
        #         os.remove(str(filepath.parent / filepath.stem) + ".docx")
        f.close()

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