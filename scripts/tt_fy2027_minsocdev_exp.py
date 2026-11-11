import pymupdf
import pandas as pd
import numpy as np
import os
from pathlib import Path
from PIL import Image
import io
import re
import pytesseract
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

pytesseract.pytesseract.tesseract_cmd = r"E:/Tesseract_OCR/tesseract.exe"
filepath = Path.cwd().joinpath("FY2027TTBudget/Draft-Estimates-of-Expenditure-for-the-Financial-Year-2027.pdf")
tesseract_dir = Path("E:/Tesseract_OCR")
tessdata_dir = tesseract_dir/"tessdata"
os.environ["TESSDATA_PREFIX"] = tessdata_dir.as_posix()

header_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Navy
category_font = Font(name="Calibri", size=11, bold=True, color="000000")
data_font = Font(name="Calibri", size=11, color="000000")
metadata_font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
metadata_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Navy

def parse_budget_line(line: str):
    line = line.strip()
    if not line:
        return None

    # clean common typos in line
    cleaned_line = line.replace("O1", "01").replace("4)", "41", 1).replace(".", ",").replace("-", ",").replace(", ", ",").replace("‘", "").replace("}", "").replace("|", "")

    # Handle parentheses as negative numbers
    cleaned_line = re.sub(r"\(\s*([\d,]+)\s*\)", r"-\1", cleaned_line)

    subhead_num = re.match(r"^(\d+)", cleaned_line)
    subhead = subhead_num.group(1) if subhead_num else ""
    

    raw_nums = re.findall(r"(-?[\d,]{4,})", cleaned_line)
    # print(raw_nums)
    # breakpoint()
    numbers = []
    for num in raw_nums:
        #clean_num = num.replace(" ", "").strip(",")
        # if clean_num and any(c.isdigit() for c in clean_num):
        if num and any(c.isdigit() for c in num):
            # numbers.append(clean_num)
            numbers.append(num.strip())

    
    if len(numbers) < 5 and len(numbers) > 0:
        # This is the subhead with a "-" in 2026 Revised Estimates
        # Add "-" to index 2
        numbers.insert(2, "-")
    elif len(numbers) == 0:
        numbers = [np.nan] * 5
    # print(numbers)
    # breakpoint()

    description = cleaned_line.replace(subhead, "", 1)
    for num in raw_nums:
        description = description.replace(num, "")
    description = re.sub(r"\s+", " ", description).strip()
    description = description.replace("Devel opment", "Development")

    subhead = subhead.replace("9", "09")
    return {
        "Sub-Head": subhead,
        "Description": description,
        "2025 Actual": numbers[0],
        "2026 Estimates": numbers[1],
        "2026 Revised Estimates": numbers[2],
        "2027 Estimates": numbers[3],
        "Variance": numbers[4]
    }



#Open Draft Estimates of Expenditure document
doc_file = pymupdf.open(filename= filepath)

expenditure_page_index = 363

raw_text_block = []
page = doc_file.load_page(expenditure_page_index)
pic = page.get_pixmap(matrix= pymupdf.Matrix(2, 2))
image_bytes = pic.tobytes("png")
image = Image.open(io.BytesIO(image_bytes))

custom_config = r"--psm 6"
text = pytesseract.image_to_string(image, config=custom_config, lang="eng")

text_list = text.split("\n")
text_subheads_list = text_list[9:]


parsed_records = []
for item  in text_subheads_list:
    record = parse_budget_line(item)
    if record:
        parsed_records.append(record)

expenditure_df = pd.DataFrame(parsed_records)
for col in expenditure_df.columns[2:]:
    expenditure_df[col] = pd.to_numeric(expenditure_df[col].astype(str).str.replace(",", "", regex=False), errors = "coerce")

expenditure_df.to_csv("TT_Ministry_of_Social_Dev.csv", index=False)

# ==========================================
# 📊 THE EXCEL CONFIGURATION INTERFACE
# ==========================================
output_path = Path.cwd()
output_file = output_path.joinpath("TT_Ministry_of_Social_Dev_Tot_Exp.xlsx")

with pd.ExcelWriter(output_file, engine= "openpyxl") as writer:
    wb = writer.book
    sheet_title = "Head_78_Expenditure_Sum"
    ws = wb.create_sheet(title= sheet_title, index=0)
    writer.sheets[sheet_title] = ws

    ws["A1"] = "HEAD: 78"
    ws["A2"] = "MINISTRY OF THE PEOPLE, SOCIAL DEVELOPMENT AND FAMILY SERVICES"

    expenditure_df.to_excel(writer, sheet_name=sheet_title, startrow=3, index=False)

    currency_format = "$#,##0";"($#,##0)";""
    header_row = 4
    start_data_row = 5
    metadata_row = [1, 2]
    end_data_row = start_data_row + expenditure_df.shape[0] - 1

    for row in range(start_data_row, end_data_row + 1):
        for col in range(3, 8):
            cell = ws.cell(row= row, column=col)
            cell.number_format = currency_format
            cell.font = data_font

    for col in range(1,8):
        cell = ws.cell(row=header_row, column=col)
        cell.font = header_font
        cell.fill = header_fill

    for row in metadata_row:
        cell = ws.cell(row= row, column= 1)
        cell.font = metadata_font
        cell.fill = metadata_fill

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_idx= col[0].column)

        for cell in col:
            if cell.value is not None:
                val_str = str(cell.value)
                if len(val_str) > max_len:
                    max_len = len(val_str)

        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    



