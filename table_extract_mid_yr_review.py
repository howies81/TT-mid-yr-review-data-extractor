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
def budget_amt_str_to_float(amount_str: str):
    if pd.isna(amount_str) or not amount_str:
        return 0

    cleaned_amt=amount_str.replace("$", "").replace(",", "").strip().lower()

    if "million" in cleaned_amt:
        num_part = cleaned_amt.replace("million", "").strip()
        num_part = float(num_part) * pow(10, 6)
        return num_part
    elif "billion" in cleaned_amt:
        num_part = cleaned_amt.replace("billion", "").strip()
        num_part = float(num_part) * pow(10, 9)
        return num_part
    else:
        return float(cleaned_amt)

        
with pdfplumber.open("Mid-Year-Review-2026-1-1.pdf") as pdf:
    recurrent_head_rows = []
    for page_num in range(43, 52):

        target_page = pdf.pages[page_num]

        target_page_text = target_page.extract_text()

        
        multiple_line_list = []
        target_page_text_split = target_page_text.split("\n")
        for line in target_page_text_split:
            if line.startswith("Head") and (len(multiple_line_list) == 0):
                match = re.search(r"Head\s+(\d{2,3}):\s+([^-–]+)[-–]\s*[\w\s]*(\$[\d]+\.*[\d]*\s+[a-zA-Z]+)", line)

                if match:
                    head_num = match.group(1)
                    entity = match.group(2).strip()
                    amount = match.group(3).strip()

                    recurrent_head_rows.append(
                        {
                            "Head Number": head_num,
                            "Entity": entity,
                            "Budget amount": amount
                        }
                    )
                else:
                    multiple_line_list.append(line)
            elif len(multiple_line_list) > 0:
                # Concatenate lines together
                line_sum = ""
                for i in range(len(multiple_line_list)):
                    line_sum = line_sum + " " + multiple_line_list[i]
                line_sum = line_sum + " " + line

                # Check to see if there is a match in the concatenated string
                match = re.search(r"Head\s+(\d{2,3}):\s+([^-–]+)[-–]\s*[\w\s]*(\$[\d]+\.*[\d]*\s+[a-zA-Z]+)", line_sum)

                if match:
                    head_num = match.group(1)
                    entity = match.group(2).strip()
                    amount = match.group(3).strip()

                    recurrent_head_rows.append(
                        {
                            "Head Number": head_num,
                            "Entity": entity,
                            "Budget amount": amount
                        }
                    )
                    multiple_line_list = []
                else:
                    multiple_line_list.append(line)

    df_recurrent = pd.DataFrame(recurrent_head_rows)
    df_recurrent["Budget amount"] = df_recurrent["Budget amount"].apply(budget_amt_str_to_float)

    development_head_rows = []
    for page_num in range(52, 54):

        target_page = pdf.pages[page_num]

        target_page_text = target_page.extract_text()

        
        multiple_line_list = []
        target_page_text_split = target_page_text.split("\n")
        for line in target_page_text_split:
            if line.startswith("Head") and (len(multiple_line_list) == 0):
                match = re.search(r"Head\s+(\d{2,3}):\s+([^-–]+)[-–]\s*[\w\s]*(\$[\d]+\.*[\d]*\s+[a-zA-Z]+)", line)

                if match:
                    head_num = match.group(1)
                    entity = match.group(2).strip()
                    amount = match.group(3).strip()

                    development_head_rows.append(
                        {
                            "Head Number": head_num,
                            "Entity": entity,
                            "Budget amount": amount
                        }
                    )
                else:
                    multiple_line_list.append(line)
            elif len(multiple_line_list) > 0:
                # Concatenate lines together
                line_sum = ""
                for i in range(len(multiple_line_list)):
                    line_sum = line_sum + " " + multiple_line_list[i]
                line_sum = line_sum + " " + line

                # Check to see if there is a match in the concatenated string
                match = re.search(r"Head\s+(\d{2,3}):\s+([^-–]+)[-–]\s*[\w\s]*(\$[\d]+\.*[\d]*\s+[a-zA-Z]+)", line_sum)

                if match:
                    head_num = match.group(1)
                    entity = match.group(2).strip()
                    amount = match.group(3).strip()

                    development_head_rows.append(
                        {
                            "Head Number": head_num,
                            "Entity": entity,
                            "Budget amount": amount
                        }
                    )
                    multiple_line_list = []
                else:
                    multiple_line_list.append(line)

    df_development = pd.DataFrame(development_head_rows)
    df_development["Budget amount"] = df_development["Budget amount"].apply(budget_amt_str_to_float)


output_file = "Trinidad_and_Tobago_2026_Mid_Year_review_summary.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as xlwriter:

    df_recurrent.to_excel(xlwriter, sheet_name="Recurrent Expenditure", index=False)

    df_development.to_excel(xlwriter, sheet_name="Development Expenditure", index=False)

