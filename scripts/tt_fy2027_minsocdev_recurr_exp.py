import pymupdf
import pandas as pd
import numpy as np
import os
from pathlib import Path
from PIL import Image, ImageEnhance, ImageFilter
import io
import re
import pytesseract
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension

pytesseract.pytesseract.tesseract_cmd = r"E:/Tesseract_OCR/tesseract.exe"
filepath = Path.cwd().joinpath("FY2027TTBudget/PRINT-VERSION-Recurrent-Expenditure-2027.pdf")
tesseract_dir = Path("E:/Tesseract_OCR")
tessdata_dir = tesseract_dir/"tessdata"
os.environ["TESSDATA_PREFIX"] = tessdata_dir.as_posix()

header_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Navy
category_font = Font(name="Calibri", size=11, bold=True, color="000000")
data_font = Font(name="Calibri", size=11, color="000000")
metadata_font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
metadata_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Navy

#Open Draft Estimates of Expenditure document
doc_file = pymupdf.open(filename= filepath)

expenditure_page_index_list = [433, 434]

raw_text_block = []
for page_index in expenditure_page_index_list:
    page = doc_file.load_page(page_index)
    pic = page.get_pixmap(matrix= pymupdf.Matrix(3.5, 3.5))
    image_bytes = pic.tobytes("png")
    image = Image.open(io.BytesIO(image_bytes))

    custom_config = r"--psm 6"
    text = pytesseract.image_to_string(image, config=custom_config, lang="eng")
    print(text)
    print("\n")
