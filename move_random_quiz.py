import csv
import sys
import os
import shutil

grade = sys.argv[1]
grade_dict = {'five': '5級', 'four': '4級', 'three': '3級',
              'pre-two': '準2級', 'two': '2級', 'pre-one': '準1級'}
jp_grade = grade_dict[grade]

page_list = []
with open(f"{grade}_page.csv", newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        page_list.append([int(row[0]), int(row[1])])


for page in page_list:
    page_begin = page[0]
    page_end = page[1]

    copy_number = 20
    for idx in range(copy_number):
        en_pdf_name = f"{grade}_en_quiz_pg{page_begin}-{page_end}_{idx + 1}.pdf"
        jp_pdf_name = f"{grade}_jp_quiz_pg{page_begin}-{page_end}_{idx + 1}.pdf"

        dir_path = f"{grade}_pg{page_begin}-{page_end}/"

        try:
            if not os.path.exists(f"{dir_path}/en/{en_pdf_name}"):
                shutil.move(en_pdf_name, f"{dir_path}/en/")
        except FileNotFoundError:
            print(f"${en_pdf_name} not found")

        try:
            if not os.path.exists(f"{dir_path}/jp/{jp_pdf_name}"):
                shutil.move(jp_pdf_name, f"{dir_path}/jp/")
        except FileNotFoundError:
            print(f"${jp_pdf_name} not found")
