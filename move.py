import os
import shutil

grade = "five"

page_begin = 14
page_end = 137

for i in range(page_begin, page_end + 1, 2):
    for idx in range(5):
        en_xlsx_name = f"{grade}_en_quiz_pg{i}-{i + 1}_{idx + 1}.xlsx"
        jp_xlsx_name = f"{grade}_jp_quiz_pg{i}-{i + 1}_{idx + 1}.xlsx"
        en_pdf_name = f"{grade}_en_quiz_pg{i}-{i + 1}_{idx + 1}.pdf"
        jp_pdf_name = f"{grade}_jp_quiz_pg{i}-{i + 1}_{idx + 1}.pdf"

        dir_path = f"{grade}_pg{i}-{i + 1}/"

        try:
            shutil.move(en_xlsx_name, f"{dir_path}/en/xlsx")
        except FileNotFoundError:
            print(f"${en_xlsx_name} not found")

        try:
            shutil.move(jp_xlsx_name, f"{dir_path}/jp/xlsx")
        except FileNotFoundError:
            print(f"${jp_xlsx_name} not found")

        try:
            shutil.move(en_pdf_name, f"{dir_path}/en/pdf")
        except FileNotFoundError:
            print(f"${en_pdf_name} not found")

        try:
            shutil.move(jp_pdf_name, f"{dir_path}/jp/pdf")
        except FileNotFoundError:
            print(f"${jp_pdf_name} not found")
