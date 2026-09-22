import pdf2docx
from pdf2docx import Converter

mid_yr_pdf = "Mid-Year-Review-2026-1-1.pdf"
mid_yr_docx = "Mid-Year-Review-2026-1-1.docx"

pdf_cv = Converter(mid_yr_pdf)
pdf_cv.convert(mid_yr_docx)
pdf_cv.close()