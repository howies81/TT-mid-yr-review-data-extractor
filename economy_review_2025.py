import re
import pandas as pd
import pdfplumber

with pdfplumber.open("WEB-•-REVIEW-OF-THE-ECONOMY-2025.pdf") as pdf:
    sample_page = pdf.pages[27]
    # print(f"Page Length: {sample_page.height}")
    # print(f"Page Width: {sample_page.width}")
    # print(f"No of geomentric rects: {len(sample_page.rects)}")
    # print(f"No of geometric lines: {len(sample_page.lines)}")

    table_data = []
    sample_page_table = sample_page.extract_table()
    if sample_page_table:
        for row in sample_page_table:
            table_data.append(row)
    #---------------------------------------------------------------------
    #---- FORMAT COUNTRY / REGION LISTING -------------------------------
    #---------------------------------------------------------------------


    country_column_list = table_data[0][0]
    raw_lines = [line.strip() for line in country_column_list.split("\n") if line.strip()]
    # print(raw_lines)
    entries_list = []
    stitched_lines = []
    stitched_line = ""
    for word in raw_lines:
        if (word.startswith("Advanced") or (word.endswith("and") and not word.startswith("Ireland"))) and len(stitched_lines) == 0:
            stitched_line = stitched_line + word
            stitched_lines.append(stitched_line)
            stitched_line = ""
        elif len(stitched_lines) > 0:
            for i in range(len(stitched_lines)):
                stitched_line = stitched_line + " " + stitched_lines[i]
            stitched_line = stitched_line + " " + word
            entries_list.append(stitched_line.strip())
            stitched_line = ""
            stitched_lines = []
        else:
            entries_list.append(word.strip())

    entries_list_df = pd.DataFrame(entries_list)

    #---------------------------------------------------------------------
    #---- FORMAT TABLE HEADERS -------------------------------
    #---------------------------------------------------------------------

    column_headers_list = table_data[0][1:]
    intermediate_headers_list =[]
    for entry in column_headers_list:
        if entry != None:
            string_entry = str(entry)
            intermediate_headers_list.append(string_entry)
        

    intermediate_headers_list_2 =[]
    for entry in intermediate_headers_list:
        split_entry = str(entry)
        split_entry = split_entry.replace("\n", " ")
        if split_entry[-1].isnumeric():
            split_entry = split_entry.replace(split_entry[-1], "")
        intermediate_headers_list_2.append(split_entry)

    years_set_list = table_data[1][1:]
    years_set_list = list(sorted(set(years_set_list)))
    print(years_set_list)
    cleaned_headers = []
    cleaned_header = ""
    for heading in intermediate_headers_list_2:
        for year in years_set_list:
            cleaned_header = heading + " " + year
            cleaned_headers.append(cleaned_header)
    print(cleaned_headers)

    #---------------------------------------------------------------------
    #---- FORMAT DATA LISTING -------------------------------
    #---------------------------------------------------------------------

    cleaned_data_rows = []

    for index, data_row in enumerate(table_data[2:]):
        cleaned_data = data_row[1:]
        for i in range(len(cleaned_data)):
            if "–" in cleaned_data[i]:
                cleaned_data[i] = cleaned_data[i].replace("–", "-")
            elif "n/a" in cleaned_data[i]:
                cleaned_data[i] = float('nan')
            else:
                cleaned_data[i] = float(cleaned_data[i])
        cleaned_data_rows.append(cleaned_data)


    cleaned_data_rows_df = pd.DataFrame(cleaned_data_rows, columns=cleaned_headers)
    cleaned_table_df = pd.concat([entries_list_df, cleaned_data_rows_df], axis=1)
    if cleaned_table_df.columns[0] == "0" or cleaned_table_df.columns[0] == 0:
        cleaned_table_df.rename(columns={cleaned_table_df.columns[0]: ""}, inplace=True)
    cleaned_table_df.to_csv("cleaned_gdp_table.csv", index=False)
    print(cleaned_table_df.to_string(index=False))
    
    

        
    
    
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

