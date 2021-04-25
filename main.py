#ファイルの読み込み
import xlrd

grade = "pretwo"

loc = (f"{grade}.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

en_problem_set = []
jp_problem_set = []

i = 1
while True:
  page_index = sheet.cell_value(i, 1)
  if page_index == "end":
    break
  #英単語を取ってくる場合は3,日本語の意味を取ってくる場合は5~
  #日本語の意味を複数個取ってこれるようにする
  en_word = sheet.cell_value(i, 3)
  jp_words = []

  j = 5
  while j <= 8:
    jp_word = sheet.cell_value(i, j)
    #print(jp_word)
    if jp_word == "":
      break

    #ここで日本語文字列から英単語を抽出して消す(予定)
    #to, with とかの言葉はリストで管理して消さないようにする

    jp_word = ''.join(jp_word.splitlines())
    jp_words.append(jp_word)
    j += 1

  en_problem_set.append([int(page_index), en_word])
  jp_problem_set.append([int(page_index), jp_words])
  i += 1

#print(en_problem_set) 
#problem_set [問題のページ番号, 単語（日本語だったら複数）]

total_length = len(en_problem_set)
current_page = en_problem_set[0][0]



### ランダムに問題を入れ替えて、ファイルに出力 ###
import random

#xlsxに書き込みできるライブラリ
import xlsxwriter
import os

idx = 0
total_problem_set = 0
while idx < total_length:
  en_problems = []
  jp_problems = []
  while idx < total_length and en_problem_set[idx][0] <= current_page + 1:

    #ページ番号２個ずつ問題セットに追加
    #英語
    en_problems.append(en_problem_set[idx][1])
    #日本語　
    jp_problems.append(jp_problem_set[idx][1])

    idx += 1


  page_begin = current_page
  page_end = current_page + 1

  #問題ページに対応したディレクトリを作成
  dir_path = (f"./{grade}_pg{page_begin}-{page_end}")
  dir_path = (f"./{grade}_pg{page_begin}-{page_end}")


  os.mkdir(dir_path)
  os.mkdir(f"{dir_path}/en")
  os.mkdir(f"{dir_path}/jp")

  os.mkdir(f"{dir_path}/en/pdf")
  os.mkdir(f"{dir_path}/en/xlsx")

  os.mkdir(f"{dir_path}/jp/pdf")
  os.mkdir(f"{dir_path}/jp/xlsx")

  #何部コピーを作るか
  copy_number = 5

  for i in range(copy_number):
    random.shuffle(en_problems)
    random.shuffle(jp_problems)

    #対応するディレクトリへxlsxファイルへの書き込み
    en_workbook = xlsxwriter.Workbook(f"{grade}_en_quiz_pg{page_begin}-{page_end}_{i + 1}.xlsx")
    en_worksheet = en_workbook.add_worksheet()
    en_cell_format = en_workbook.add_format()

    jp_workbook = xlsxwriter.Workbook(f"{grade}_jp_quiz_pg{page_begin}-{page_end}_{i + 1}.xlsx")
    jp_worksheet = jp_workbook.add_worksheet()
    jp_cell_format = jp_workbook.add_format()

    en_pg_cell_format = en_workbook.add_format()
    jp_pg_cell_format = jp_workbook.add_format()


    #Cellのフォーマット指定

    en_pg_cell_format.set_font_size(16)
    jp_pg_cell_format.set_font_size(16)
    en_pg_cell_format.set_font_name('Times New Roman')
    jp_pg_cell_format.set_font_name('Times New Roman')
    en_pg_cell_format.set_align('center')
    jp_pg_cell_format.set_align('center')
    en_pg_cell_format.set_align('vcenter')
    jp_pg_cell_format.set_align('vcenter')
    


    #英語
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

    #日本語


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



    #15 * 4のテーブル
    length = len(en_problems)
    mod = 9

    #英語の問題
    en_worksheet.set_header(f'&C&16&"Hiragino Sans,Regular"英検準2級単語テスト&R&16&"Times New Roman,Regular"', {'margin':0.5})

    en_worksheet.write(0, 3, f'Pg{page_begin}-{page_end}', en_pg_cell_format)

    #en_worksheet.set_footer('&R&G', {'image_right': 'LOGO_B1_copy1.jpg'})
    en_worksheet.set_margins(left=0.7, right = 0.0, top = 0.75, bottom = 0.75)
    en_worksheet.insert_image('C12', 'LOGO_B1.jpg', {'x_offset':76.5, 'y_offset':30, 'x_scale': 0.045, 'y_scale': 0.045})
    en_worksheet.center_horizontally()

    row_base = 1
    col_base = 0
    row = 0
    col = 0
    add = 0
    index = 1

    for i in range(18): 
      problem = ""
      if i < length:
        problem = en_problems[i]

      en_worksheet.set_row(row, 60)
      en_worksheet.write((row % mod) + row_base, col + col_base + add, problem, en_cell_format)
      en_worksheet.write((row % mod) + row_base, col + 1 + col_base + add, "", en_cell_format)

      row += 1
      index += 1
      if mod <= row and add == 0:
        add = 2


    en_workbook.close()

    #日本語の問題
    jp_worksheet.set_header(f'&C&16&"Hiragino Sans,Regular"英検準2級単語テスト&R&16&"Times New Roman,Regular"', {'margin':0.5})
    
    jp_worksheet.write(0, 3, f'Pg{page_begin}-{page_end}', jp_pg_cell_format)

    jp_worksheet.set_margins(left=0.7, right = 0.0, top = 0.75, bottom = 0.75)
    #jp_worksheet.set_footer('&R&G', {'image_right': 'LOGO_B1_copy1.jpg'})
    jp_worksheet.insert_image('C12', 'LOGO_B1.jpg', {'x_offset':76.5, 'y_offset':30, 'x_scale': 0.045, 'y_scale': 0.045})
    jp_worksheet.center_horizontally()

    row_base = 1
    col_base = 0
    row = 0
    col = 0
    add = 0
    index = 1

    for i in range(18):
      problem = ""

      if i < length:
        problem_list = jp_problems[i]
        problem_length = len(problem_list)

        #for meaning in problem_list:
        for j in range(problem_length):
          meaning = problem_list[j]
          if j == problem_length - 1:
            problem += meaning
          else:
            problem += meaning + '、'

      jp_worksheet.set_row(row, 60)
      jp_worksheet.write((row % mod) + row_base, col + col_base + add, problem, jp_cell_format)
      jp_worksheet.write((row % mod) + row_base, col + 1 + col_base + add, "", jp_cell_format)
        

      row += 1
      index += 1
      if mod <= row and add == 0:
        add = 2
      
    
    jp_workbook.close()

  current_page = current_page + 2
  
  total_problem_set += 1

