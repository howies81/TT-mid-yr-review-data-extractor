import pandas as pd
import re
import numpy as np
import pdfplumber
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

filename = "WEB-•-REVIEW-OF-THE-ECONOMY-2025.pdf"
LINE_TOL = 3.0
header_font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # Classic Navy
category_font = Font(name="Calibri", size=11, bold=True, color="000000")
data_font = Font(name="Calibri", size=11, color="000000")

def extract_revenue_table(filename: str, page_index: int) -> pd.DataFrame | None:
    try:
        with pdfplumber.open(filename) as pdf:
            sample_page = pdf.pages[page_index]
            word_bank = sample_page.extract_words()
            word_bank_df = pd.DataFrame(word_bank)
        return word_bank_df

    except FileExistsError or FileNotFoundError:
        print("File Not Found!")
        return None

def row_aggregator(word_bank_df: pd.DataFrame, tolerance: float = LINE_TOL) -> list[pd.DataFrame] | None:
    try:
        if word_bank_df.empty:
            return None
        row_word_df = word_bank_df.sort_values(by= "top")
        row_id = (row_word_df["top"].diff() > tolerance).cumsum()
        sorted_row_list = []
        for _, g in row_word_df.groupby(row_id):
            sorted_row_list.append(g.sort_values(by="x0"))
        return sorted_row_list

    except Exception as e:
        print(f"Row aggregation error {e}")
        return None

def convert_to_number(cell: str):
    cell = str(cell).strip()

    if cell in ("", "-", "N/A"):
        return np.nan
    try:
        value = float(cell.strip("()").replace(",", ""))
    except ValueError:
        return np.nan
    return value

def parse_revenue_table(filename: str, page_index: int) -> pd.DataFrame | None:
    df = extract_revenue_table(filename=filename, page_index=page_index)
    if df is None:
        return None

    
    # Aggregate rows on page
    page_row_df = row_aggregator(word_bank_df=df)
    if page_row_df is None:
        return None

    #Define table boundaries
    header_top = df.loc[df["text"] == "Oct", "top"].min()
    header_btm = df.loc[df["text"] == "Sep", "bottom"].min()
    source_top = df.loc[df["text"].str.startswith("Source"), "top"].min()

    # Header top = 135.67
    # Header bottom = 156.32
    # Source top = 723
    
    # Therefore table header lies between y = 135.76 and y = 156.32
    # Table with labels and data lies between y = 156.32 and y = 723

    # Extract header text 
    header_text = df[(df["top"] > header_top - 4) & (df["top"] < header_btm + 3)]
    header_list = row_aggregator(word_bank_df=header_text)
    if header_list is None:
        return None

    # Extract body table from header row
    body_table = df[(df["top"] >= header_btm + 3) & (df["top"] < source_top - 4)]
    if len(body_table) == 0:
        return None

    # Establish separation line between labels and data
    # 
    target_row = body_table.iloc[0]
    if target_row is None:
        return None
    numeric_pattern = r"^[0-9\.,\(\)\-]+$"
    is_numeric = body_table["text"].str.match(numeric_pattern, na=False)
    numeric_rows = body_table.loc[is_numeric]
    if not numeric_rows.empty:
        split_x_table = (numeric_rows["x0"].iloc[0]) - 4
    else:
        split_x_table = body_table["x0"].min()
    
        
    # Split body table into labels and data
    label_column = body_table[body_table["x0"] < split_x_table]
    data_table = body_table[body_table["x0"] >= split_x_table]

    # Create header column labels
    col_boundaries = []
    for i in range(6):
        boundary = (data_table["x0"].iloc[i]) - 4
        col_boundaries.append(boundary)

    # Rearrange header row into 6 columns
    restructured_rows = []
    for row in header_list:
        
        col_slots = [""] * len(col_boundaries)

        for index, word in row.iterrows():
            
            # Find the closest matching column boundary using a simple index lookup
            # It finds the last boundary that is smaller than or equal to the word's x0
            col_index = max([i for i, b in enumerate(col_boundaries) if word["x0"] >= b], default=0)
            if col_slots[col_index] == "":
                col_slots[col_index] = word["text"]
            else:
                col_slots[col_index] += " " + word["text"]
        restructured_rows.append(col_slots)
    new_header_df = pd.DataFrame(restructured_rows)
    new_header_df_int = new_header_df.agg(lambda x: " ".join(x.astype(str))
    )
    new_header_df = pd.DataFrame([new_header_df_int])
    new_header_list = list(new_header_df_int)
    

    #Neext is to aggregate the rows in the label column so that the text lines up like in the
    # Appendix page
    
    label_rows = row_aggregator(label_column)
    if label_rows is None:
        return None

    
    restructured_label_rows = []
    for row in label_rows:
        
        row_text = ""
        for index, word in row.iterrows():
            row_text += " " + word["text"]
        restructured_label_rows.append(row_text.strip())
    # 41 rows for Revenue sheet, 
    label_column_df = pd.DataFrame(restructured_label_rows, columns=[""])
    

    no_data_rows = []
    for seq_idx, (index, row) in enumerate(label_column_df.iterrows()):
        
        if ("of which:-" in row.values) or ("of which: -" in row.values):
            no_data_rows.append(seq_idx)
    # no_data_rows = [5, 17, 22, 26, 37]
    


    
    

    data_rows = row_aggregator(data_table)
    
    if data_rows is None:
        return None
    

    nan_list = [np.nan] * 6
    
    restructured_data_rows = []

    
    for curr_row_index, row in enumerate(data_rows):

        word_list = [""] * len(col_boundaries)

        for index, word in row.iterrows():
            x_pos = word["x0"]
            w_text = word["text"]

            col_index = max([i for i, b in enumerate(col_boundaries) if x_pos >= b], default=0)

            if word_list[col_index] == "":
                word_list[col_index] = w_text
            else:
                word_list[col_index] += " " + w_text

        word_list = [np.nan if val == "" else val for val in word_list]
        restructured_data_rows.append(word_list)

    for row_no in no_data_rows:
        restructured_data_rows.insert(row_no, nan_list)
        
    
    data_table_df = pd.DataFrame(restructured_data_rows, columns=new_header_list)
    for col in new_header_list:
        data_table_df[col] = data_table_df[col].map(convert_to_number)
    full_table_df = pd.concat([label_column_df, data_table_df], axis=1)
    
    
    return full_table_df
    
        

if __name__ == "__main__":

    with pd.ExcelWriter("Central_Gov_Revenue_TT.xlsx", engine="openpyxl") as writer:

        df = parse_revenue_table(filename=filename, page_index=210)
        if df is not None:
            sheet_tab_name = "Central_GovTT_Revenue"
            df.to_excel(writer, sheet_name=sheet_tab_name, index=False)
            num_format = "#,##0.0"
            active_sheet = writer.sheets[sheet_tab_name]

            #Change number format
            for row in active_sheet.iter_rows(min_row= 2, min_col= 2):
                for cell in row:
                    cell.number_format = num_format
                    cell.font = data_font

            #Format header
            for cell in active_sheet[1]:
                cell.font = header_font
                cell.fill = header_fill

            # Format table similar to PDF table
            bold_rows = [2, 3, 4, 5, 6, 17, 18, 23, 26, 27, 38]
            for row_num in bold_rows:
                for cell in active_sheet[row_num]:
                    cell.font = Font(bold=True)

            for col in active_sheet.columns:
                max_len = 0
                col_letter = col[0].column_letter

                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        if len(val_str) > max_len:
                            max_len = len(val_str)

                active_sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
                    