import csv
import sys
import os
import random
import xlsxwriter
import xlrd

grade = sys.argv[1]
grade_dict = {'five': '5級', 'four': '4級', 'three': '3級',
              'pre-two': '準2級', 'two': '2級', 'pre-one': '準1級'}
jp_grade = grade_dict[grade]


loc = (f"{grade}.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

data = []

i = 1
while True:
    try:
        row = sheet.row_values(i)
        data.append(row)
    except IndexError:
        break
    i = i + 1

page_list = []
with open(f"{grade}_page.csv", newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        page_list.append([int(row[0]), int(row[1])])


for page in page_list:
    page_begin = page[0]
    page_end = page[1]

    en_problem_candidates = []
    jp_problem_candidates = []

    for row in data:
        try:
            current_page = int(row[1])
        except ValueError:
            break

        if not page_begin <= current_page <= page_end:
            continue

        jp_idx = 5
        en_word = str(row[3])
        jp_words = []
        while True:
            try:
                jp_word = row[jp_idx]
                if len(jp_word) == 0:
                    break

                jp_word = ''.join(str(jp_word).splitlines())
                jp_words.append(jp_word)
                jp_idx += 1
            except IndexError:
                break

        en_problem_candidates.append(en_word)
        jp_problem_candidates.append(jp_words)

    problem_number = 20

    # 問題ページに対応したディレクトリを作成
    dir_path = (f"./{grade}_pg{page_begin}-{page_end}")

    if not os.path.exists(dir_path):
        os.mkdir(dir_path)

    if not os.path.exists(f"{dir_path}/en"):
        os.mkdir(f"{dir_path}/en")

    if not os.path.exists(f"{dir_path}/jp"):
        os.mkdir(f"{dir_path}/jp")

    # 何部コピーを作るか
    copy_number = 5
    for i in range(copy_number):
        en_problems = random.sample(en_problem_candidates, problem_number)
        jp_problems = random.sample(jp_problem_candidates, problem_number)

        # 対応するディレクトリへxlsxファイルへの書き込み
        en_workbook = xlsxwriter.Workbook(
            f"{grade}_en_quiz_pg{page_begin}-{page_end}_{i + 1}.xlsx")
        en_worksheet = en_workbook.add_worksheet()
        en_cell_format = en_workbook.add_format()

        jp_workbook = xlsxwriter.Workbook(
            f"{grade}_jp_quiz_pg{page_begin}-{page_end}_{i + 1}.xlsx")
        jp_worksheet = jp_workbook.add_worksheet()
        jp_cell_format = jp_workbook.add_format()

        en_pg_cell_format = en_workbook.add_format()
        jp_pg_cell_format = jp_workbook.add_format()

        # Cellのフォーマット指定

        en_pg_cell_format.set_font_size(16)
        jp_pg_cell_format.set_font_size(16)
        en_pg_cell_format.set_font_name('Times New Roman')
        jp_pg_cell_format.set_font_name('Times New Roman')
        en_pg_cell_format.set_align('center')
        jp_pg_cell_format.set_align('center')
        en_pg_cell_format.set_align('vcenter')
        jp_pg_cell_format.set_align('vcenter')

        # 英語
        en_cell_format.set_text_wrap()
        en_cell_format.set_font_size(15)
        en_cell_format.set_align('top')
        en_cell_format.set_align('vcenter')
        en_cell_format.set_border()
        en_cell_format.set_font_name('Times New Roman')

        en_worksheet.set_column('A:A', width=14)
        en_worksheet.set_column('B:B', width=24)
        en_worksheet.set_column('C:C', width=14)
        en_worksheet.set_column('D:D', width=24)

        # 日本語
        jp_cell_format.set_text_wrap()
        jp_cell_format.set_font_size(10)
        jp_cell_format.set_align('top')
        jp_cell_format.set_align('vcenter')
        jp_cell_format.set_border()
        jp_cell_format.set_shrink()
        jp_cell_format.set_font_name('Hiragino Sans')

        jp_worksheet.set_column('A:A', width=22)
        jp_worksheet.set_column('B:B', width=16)
        jp_worksheet.set_column('C:C', width=22)
        jp_worksheet.set_column('D:D', width=16)

        # 15 * 4のテーブル
        length = len(en_problems)
        mod = 10

        # 英語の問題
        en_worksheet.set_header(
            f'&C&16&"Hiragino Sans,Regular"英検{jp_grade}単語まとめテスト&R&16&"Times New Roman,Regular"', {'margin': 0.5})

        en_worksheet.write(
            0, 3, f'Pg{page_begin}-{page_end}', en_pg_cell_format)

        en_worksheet.set_margins(left=0.7, right=0.0, top=0.75, bottom=0.75)
        en_worksheet.insert_image('C12', 'LOGO_B1.jpg', {
                                  'x_offset': 76.5, 'y_offset': 30, 'x_scale': 0.045, 'y_scale': 0.045})
        en_worksheet.center_horizontally()

        # What am I doing here?
        row_base = 1
        col_base = 0
        row = 0
        col = 0
        add = 0
        index = 1

        for i in range(problem_number):
            problem = ""
            if i < length:
                problem = en_problems[i]

            en_worksheet.set_row(row, 60)
            en_worksheet.write((row % mod) + row_base, col +
                               col_base + add, problem, en_cell_format)
            en_worksheet.write((row % mod) + row_base, col +
                               1 + col_base + add, "", en_cell_format)

            row += 1
            index += 1
            if mod <= row and add == 0:
                add = 2

        en_workbook.close()

        # 日本語の問題
        jp_worksheet.set_header(
            f'&C&16&"Hiragino Sans,Regular"英検{jp_grade}単語まとめテスト&R&16&"Times New Roman,Regular"', {'margin': 0.5})

        jp_worksheet.write(
            0, 3, f'Pg{page_begin}-{page_end}', jp_pg_cell_format)

        jp_worksheet.set_margins(left=0.7, right=0.0, top=0.75, bottom=0.75)
        jp_worksheet.insert_image('C12', 'LOGO_B1.jpg', {
                                  'x_offset': 76.5, 'y_offset': 30, 'x_scale': 0.045, 'y_scale': 0.045})
        jp_worksheet.center_horizontally()

        row_base = 1
        col_base = 0
        row = 0
        col = 0
        add = 0
        index = 1

        for i in range(problem_number):
            problem = ""

            if i < length:
                problem_list = jp_problems[i]
                problem_length = len(problem_list)

                # for meaning in problem_list:
                for j in range(problem_length):
                    meaning = problem_list[j]
                    if j == problem_length - 1:
                        problem += meaning
                    else:
                        problem += meaning + '、'

            jp_worksheet.set_row(row, 60)
            jp_worksheet.write((row % mod) + row_base, col +
                               col_base + add, problem, jp_cell_format)
            jp_worksheet.write((row % mod) + row_base, col +
                               1 + col_base + add, "", jp_cell_format)

            row += 1
            index += 1
            if mod <= row and add == 0:
                add = 2

        jp_workbook.close()
