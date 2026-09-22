import pdfplumber
import pandas as pd
import re

# -----------------------------------------------------------
# COORDINATE MAPPING SCRIPT
# -----------------------------------------------------------

# with pdfplumber.open("Mid-Year-Review-2026-1-1.pdf") as pdf:
#     # for i, page in enumerate(pdf.pages):
#     #     text = page.extract_text()
#     # if text:
#     #     print(text[:500])

#     target_page = pdf.pages[43]
#     target_page_width = target_page.width
#     target_page_height = target_page.height
#     print(f"Page Width: {target_page_width}")
#     print(f"Page Height: {target_page_height}")

#     page_heads = []
#     page_words = target_page.extract_words()
#     for word in page_words:
#         if "Head" in word["text"]:
#             page_heads.append(
#                 {
#                     "Text": word["text"],
#                     "Left Edge (x0)": word["x0"],
#                     "Right Edge (x1)": word["x1"],
#                     "Top Edge (top)": word["top"],
#                     "Bottom Edge (bottom)": word["bottom"],
#                     "Doctop Value (doctop)": word["doctop"]

#                 }
#             )
#     df = pd.DataFrame(page_heads)
#     print("---Geometric positions of 'Head' on page 44")
#     print(df.to_string(index=False))
        
with pdfplumber.open("Mid-Year-Review-2026-1-1.pdf") as pdf:
    target_page = pdf.pages[1]
    target_page_width = target_page.width
    
    vertical_borders = [70, 518, 540]

    table_settings = {
        "vertical_strategy": "explicit",
        "explicit_vertical_lines": vertical_borders,
        "horizontal_strategy": "text",
        "snap_y_tolerance": 4
    }

    raw_matrix = target_page.extract_table(table_settings=table_settings)
    

    if raw_matrix:
        cleaned_rows = []
        stitched_rows = []
        for row in raw_matrix:
            if not row[0] and not row[1]:
                continue

            section_title = re.sub(r"\.\.+", "", row[0]).strip() if row[0] else ""
            page_num = row[1].lstrip(" .") if row[1] else ""
            if section_title and section_title != 'Contents' and page_num and len(stitched_rows) == 0:
                cleaned_rows.append(
                    {
                        "Section Title": section_title,
                        "Page Number": page_num
                    }
                )
            elif section_title and section_title != 'Contents' and not page_num:
                stitched_rows.append(section_title)
                
            elif section_title and section_title != 'Contents' and page_num and len(stitched_rows) > 0:
                line_sum = ""
                for i in range(len(stitched_rows)):
                    line_sum = line_sum + " " + stitched_rows[i]
                line_sum = line_sum + " " + section_title
                cleaned_rows.append({
                    "Section Title": line_sum,
                    "Page Number": page_num
                })
                stitched_rows = []


    df_toc = pd.DataFrame(cleaned_rows)
    df_toc["Page Number"] = pd.to_numeric(df_toc["Page Number"], errors="coerce").astype("Int64")
    print(df_toc.to_string(index=False))

    df_toc.to_csv("Mid_Year_Review_TOC.csv", index=False)



