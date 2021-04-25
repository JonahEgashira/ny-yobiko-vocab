'''
構想
ディレクトリを探索する
pgの判断をファイル名からする->末尾の数字?（桁数が違うのをどう処理するか)
pgに対して動詞Aとか形容詞Bとかのフォルダで分けるのは手動でページ数の範囲を打ち込んで処理する
移動自体はshutil.moveで楽にできそう

2級
|-動詞A
| |- en10-11
| |- jp10-11
| |- ...
|-名詞A
| |- en24-25
| |- jp24-25
|


'''

import os, shutil


def get_section(page, grade):
    '''
    終わりのページに対しての範囲を返す
    '''
    if grade == "two":
        if 10 <= page <= 27:
            return "動詞A"
        elif 28 <= page <= 45:
            return "名詞A"
        elif 46 <= page <= 55:
            return "形容詞・副詞・その他A"
        elif 60 <= page <= 77:
            return "動詞B"
        elif 78 <= page <= 95:
            return "名詞B"
        elif 96 <= page <= 105:
            return "形容詞・副詞・その他B"
        elif 110 <= page <= 125:
            return "動詞C"
        elif 126 <= page <= 143:
            return "名詞C"
        elif 144 <= page <= 151:
            return "形容詞・副詞・その他C"
        elif 156 <= page <= 193:
            return "熟語A"
        elif 198 <= page <= 235:
            return "熟語B"
        else:
            return "None"
    
    if grade == "pre_two":
        if 10 <= page <= 23:
            return "動詞A"
        elif 24 <= page <= 39:
            return "名詞A"
        elif 40 <= page <= 51:
            return "形容詞・副詞・その他A"
        elif 56 <= page <= 68:
            return "動詞B"
        elif 69 <= page <= 85:
            return "名詞B"
        elif 86 <= page <= 97:
            return "形容詞・副詞・その他B"
        elif 102 <= page <= 113:
            return "動詞C"
        elif 114 <= page <= 130:
            return "名詞C"
        elif 131 <= page <= 141:
            return "形容詞・副詞・その他C"
        elif 146 <= page <= 175:
            return "熟語A"
        elif 180 <= page <= 209:
            return "熟語B"
        else:
            return "None"


def get_page(filename):
    last_page = 0
    power = 1
    start = False
    for x in reversed(filename):
        if x == '_':
            start = True
            continue
        if x == '-':
            break
        if not start:
            continue

        last_page += int(x) * power
        power *= 10
    return last_page


def get_lang(file_name):
    return file_name[4:6]

path_from = "two_from"
path_to = "two_to"

def main(path_from, path_to, grade):
    os.mkdir(f"{path_to}/pdf")
    os.mkdir(f"{path_to}/xlsx")

    used_set = set()
    section_set = set()
    for pathname, _, filenames in os.walk(path_from):
        for filename in filenames:
            file_type = ""
            if filename[-3:] == "pdf":
                file_type = "pdf"
            elif filename[-4:] == "xlsx":
                file_type = "xlsx"
            else:
                continue

            last_page = get_page(filename)
            language = get_lang(filename)
            section = get_section(last_page, grade)

            info = language + str(last_page) + file_type
            if info in used_set:
                continue

            used_set.add(info)

            if section == "None":
                continue

            section_type = file_type+section
            if section_type not in section_set:
                os.mkdir(f"{path_to}/{file_type}/{section}")
                section_set.add(section_type)

            new_file_name = f"{last_page-1}-{last_page}-{language}.{file_type}"

            #rename
            shutil.copy(f"{pathname}/{filename}", f"{pathname}/{new_file_name}")

            #move
            shutil.move(f"{pathname}/{new_file_name}", f"{path_to}/{file_type}/{section}")


import PyPDF2

def compress(path, file_type):
    for directory in os.listdir(path):
        if directory.startswith('.'):
            continue

        merger = PyPDF2.PdfFileMerger()
        files = os.listdir(os.path.join(path,directory))
        files.sort()
        for file in files:
            file_path = os.path.join(path, directory, file)
            merger.append(file_path)

        new_path = f"{path}/{directory}.{file_type}"
        print(new_path)
        merger.write(new_path)
        merger.close()
        print("-------------------")


compress("pre_two_to/pdf", "pdf")
#path_from = "pre_two_from"
#path_to = "pre_two_to"
#main(path_from, path_to, "pre_two")