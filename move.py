import os
import shutil

grade = "three"

page_begin = 10
page_end = 190

for i in range(page_begin, page_end + 1, 2):
  for idx in range(5):
    en_xlsx_name = f"{grade}_en_quiz_pg{i}-{i + 1}_{idx + 1}.xlsx"
    jp_xlsx_name = f"{grade}_jp_quiz_pg{i}-{i + 1}_{idx + 1}.xlsx"
    en_pdf_name = f"{grade}_en_quiz_pg{i}-{i + 1}_{idx + 1}.pdf"
    jp_pdf_name = f"{grade}_jp_quiz_pg{i}-{i + 1}_{idx + 1}.pdf"


    dir_path = f"{grade}_pg{i}-{i + 1}/"

    shutil.move(en_xlsx_name, f"{dir_path}/en/xlsx")
    shutil.move(jp_xlsx_name, f"{dir_path}/jp/xlsx")
    shutil.move(en_pdf_name, f"{dir_path}/en/pdf")
    shutil.move(jp_pdf_name, f"{dir_path}/jp/pdf")
