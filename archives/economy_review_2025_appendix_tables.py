import re
import pandas as pd
from pandas import io
import numpy as np
import pdfplumber

def check_number(text_item):
    numeric_pattern = r"^-?\d+(?:\.\d+)?$"

    return bool(re.match(numeric_pattern, str(text_item).strip()))

def clean_and_check_number(string_item: str) -> bool:
    try:
        float(str(string_item).replace(",", "").strip())
        return True
    except ValueError:
        return False

NUM_DASHES = {"-", "–", "—", "N/A", "n/a"}

def to_value(tok):
    if tok in NUM_DASHES:
        return float("nan")
    t = tok.replace(",", "").replace("–", "-").replace("−", "-")
    if t.startswith("(") and t.endswith(")"):
        t = "-" + t[1:-1]
    try:
        return float(t)
    except ValueError:
        return float("nan")

def group_lines(words, tol=3):
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(w["top"] - lines[-1][0]["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]

def parse_page(page):
    table = page.find_tables()[0]
    header_bottom = table.rows[0].bbox[3]

    # column edges for the label and ISIC columns, taken from the body row
    body_cells = [c for c in table.rows[1].cells if c is not None]
    label_right = body_cells[0][2]
    isic_right = body_cells[1][2]

    words = page.crop(table.bbox).extract_words(x_tolerance=2, y_tolerance=2)
    words = [w for w in words if w["top"] >= header_bottom - 1]

    records, pending = [], []
    for line in group_lines(words):
        label, isic, vals = [], None, []
        for w in line:
            if w["x1"] <= label_right + 1:
                label.append(w["text"])
            elif w["x1"] <= isic_right + 1:
                isic = w["text"]
            else:
                vals.append(to_value(w["text"]))

        text = re.sub(r"\s*\d{1,2}$", "", " ".join(label))  # strip footnote digits

        if not vals and isic is None:      # wrapped label, no data on this line
            pending.append(text)
            continue

        text = " ".join(pending + [text]).strip()
        pending = []
        records.append([text, isic, *vals])

    return pd.DataFrame(records)

with pd.ExcelWriter("GDP_Analysis_Output.xlsx", engine="openpyxl") as writer:
    excel_label_map = {
        188: "GDP Constant 2012 Prices TTD Mn",
        189: "GDP Constant 2012 Prices Pct Change",
        190: "GDP Constant 2012 Prices Pct Cont",
        191: "GDP Current 2012 Prices TTD Mn",
        192: "GDP Current 2012 Prices Pct Change",
        193: "GDP Current 2012 Prices Pct Cont"
    }

    # excel_label_map_df = pd.DataFrame(excel_label_map)
    
    with pdfplumber.open("WEB-•-REVIEW-OF-THE-ECONOMY-2025.pdf") as pdf:
        for page_index in range(188, 194):
            sample_page = pdf.pages[page_index]
            
            table_data = []
            sample_page_table = sample_page.extract_table()
            if sample_page_table:
                for row in sample_page_table:
                    table_data.append(row)
            
            #---------------------------------------------------------------------
            #---- FORMAT HEADER LISTING -------------------------------
            #---------------------------------------------------------------------
            suffix_map = {
                "e": "(MOF Estimated)",
                "r": "(CSO Revised)",
                "p": "(CSO Provisional)"
            }
            header_list = table_data[0]
            cleaned_header_list = []
            for item in header_list:
                if item is None:
                    #cleaned_header_list.append("")
                    continue
                item = str(item).strip()
                if item == "":
                    continue
                
                
                if " " in item:
                    item_sub_tokens = item.split()

                    token_processing = []
                    skip_next = False
                    print(len(item_sub_tokens))

                    for index in range(len(item_sub_tokens)):
                        if skip_next:
                            skip_next = False
                            continue

                        current_token = item_sub_tokens[index]

                        if re.match(r"^Q\d+\*$", current_token) and ((index + 1) < len(item_sub_tokens)):
                            fused_q_header = f"{current_token} {item_sub_tokens[index + 1]}"
                            token_processing.append(fused_q_header)
                            skip_next = True
                        else:
                            token_processing.append(current_token)
                else:
                    token_processing = [item]
                        
                for tokens in token_processing:
                    tokens.strip()
                    last_char = tokens[-1].lower()
                    if last_char in suffix_map:
                        cleaned_token = tokens[:-1] + " " + suffix_map[last_char]
                    else:
                        cleaned_token = tokens
                    cleaned_header_list.append(cleaned_token)
            
            

            df = parse_page(sample_page)
            print(page_index, df.shape, "headers:", len(cleaned_header_list))
            print(df.to_string())

            if df.shape[1] == len(cleaned_header_list):
                df.columns = cleaned_header_list
            else:
                print("MISMATCH on page", page_index)
                df.to_csv(f"debug_page_{page_index}.csv", index=False)
                continue

            sheet_tab_name = excel_label_map.get(page_index, f"Appendix_Page_{page_index + 1}")
            df.to_excel(writer, sheet_name=sheet_tab_name, index=False)
                        
            
            
            #final_extracted_table_df.to_csv("cleaned_gdp_appendix_table.csv", index=False)
            #final_extracted_table_df.to_excel(writer, sheet_name=)

    # country_column_list = table_data[0][0]
    # raw_lines = [line.strip() for line in country_column_list.split("\n") if line.strip()]
    # # print(raw_lines)
    # entries_list = []
    # stitched_lines = []
    # stitched_line = ""
    # for word in raw_lines:
    #     if (word.startswith("Advanced") or (word.endswith("and") and not word.startswith("Ireland"))) and len(stitched_lines) == 0:
    #         stitched_line = stitched_line + word
    #         stitched_lines.append(stitched_line)
    #         stitched_line = ""
    #     elif len(stitched_lines) > 0:
    #         for i in range(len(stitched_lines)):
    #             stitched_line = stitched_line + " " + stitched_lines[i]
    #         stitched_line = stitched_line + " " + word
    #         entries_list.append(stitched_line.strip())
    #         stitched_line = ""
    #         stitched_lines = []
    #     else:
    #         entries_list.append(word.strip())

    # entries_list_df = pd.DataFrame(entries_list)

    # #---------------------------------------------------------------------
    # #---- FORMAT TABLE HEADERS -------------------------------
    # #---------------------------------------------------------------------

    # column_headers_list = table_data[0][1:]
    # intermediate_headers_list =[]
    # for entry in column_headers_list:
    #     if entry != None:
    #         string_entry = str(entry)
    #         intermediate_headers_list.append(string_entry)
        

    # intermediate_headers_list_2 =[]
    # for entry in intermediate_headers_list:
    #     split_entry = str(entry)
    #     split_entry = split_entry.replace("\n", " ")
    #     if split_entry[-1].isnumeric():
    #         split_entry = split_entry.replace(split_entry[-1], "")
    #     intermediate_headers_list_2.append(split_entry)

    # years_set_list = table_data[1][1:]
    # years_set_list = list(sorted(set(years_set_list)))
    # print(years_set_list)
    # cleaned_headers = []
    # cleaned_header = ""
    # for heading in intermediate_headers_list_2:
    #     for year in years_set_list:
    #         cleaned_header = heading + " " + year
    #         cleaned_headers.append(cleaned_header)
    # print(cleaned_headers)

    # #---------------------------------------------------------------------
    # #---- FORMAT DATA LISTING -------------------------------
    # #---------------------------------------------------------------------

    # cleaned_data_rows = []

    # for index, data_row in enumerate(table_data[2:]):
    #     cleaned_data = data_row[1:]
    #     for i in range(len(cleaned_data)):
    #         if "–" in cleaned_data[i]:
    #             cleaned_data[i] = cleaned_data[i].replace("–", "-")
    #         elif "n/a" in cleaned_data[i]:
    #             cleaned_data[i] = float('nan')
    #         else:
    #             cleaned_data[i] = float(cleaned_data[i])
    #     cleaned_data_rows.append(cleaned_data)


    # cleaned_data_rows_df = pd.DataFrame(cleaned_data_rows, columns=cleaned_headers)
    # cleaned_table_df = pd.concat([entries_list_df, cleaned_data_rows_df], axis=1)
    # if cleaned_table_df.columns[0] == "0" or cleaned_table_df.columns[0] == 0:
    #     cleaned_table_df.rename(columns={cleaned_table_df.columns[0]: ""}, inplace=True)
    # cleaned_table_df.to_csv("cleaned_gdp_table.csv", index=False)
    # print(cleaned_table_df.to_string(index=False))
    
    

        
    
    
    # column_years_list = table_data[1][:]
    # print(column_years_list)
    # column_years_list = [entry.strip() for entry in column_years_list if entry != None]
    # print(column_years_list)
           

    # x_edges = set()
    # for rect in sample_page.rects:
    #     print(rect["x0"])
    #     print(rect["x1"])
    #     x_edges.add(rect["x0"])
    #     x_edges.add(rect["x1"])
    #     print(x_edges)
    #     breakpoint()

    
    # x_edges_sorted = sorted(list(x_edges))
    # filtered_fences = []
    # for edge in x_edges_sorted:
    #     if edge < 45:
    #         continue

    #     if not filtered_fences or (edge - filtered_fences[-1]) > 3:
    #         filtered_fences.append(edge)

    # print(len(filtered_fences))
    # print(filtered_fences)

    # table_settings = {
    #     "vertical_strategy": "explicit",
    #     "explicit_vertical_lines": filtered_fences,
    #     "horizontal_strategy": "text", # Group rows line-by-line horizontally
    #     "snap_y_tolerance": 3
    # }
    
    # # Extract the nested matrix
    # raw_matrix = sample_page.extract_table(table_settings=table_settings)
    # print(raw_matrix)
    # sample_table = sample_page.extract_table()
    #print(sample_table)
    # df_table = pd.DataFrame(sample_table)
    # print(df_table.to_string(index=False, header=False))
    # df_table.to_csv("sample_table.csv", index=False, header=False)

