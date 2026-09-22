import re
import pandas as pd
import pdfplumber

with pdfplumber.open("WEB-•-REVIEW-OF-THE-ECONOMY-2025.pdf") as pdf:
    sample_page = pdf.pages[45]
    sample_table = sample_page.extract_table()
    #print(sample_table)
    df_table = pd.DataFrame(sample_table)
    print(df_table.to_string(index=False, header=False))
    df_table.to_csv("sample_table.csv", index=False, header=False)

