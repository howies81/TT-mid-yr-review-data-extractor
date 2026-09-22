import re
import pandas as pd
import pdfplumber

with pdfplumber.open("WEB-•-REVIEW-OF-THE-ECONOMY-2025.pdf") as pdf:
    sample_page = pdf.pages[27]
    print(f"Page Length: {sample_page.height}")
    print(f"Page Width: {sample_page.width}")
    print(f"No of geomentric rects: {len(sample_page.rects)}")
    print(f"No of geometric lines: {len(sample_page.lines)}")

    raw_text = sample_page.extract_text() or ""
    if "Advanced Economies" in raw_text:
        print(raw_text[:9])
    print(raw_text.split("\n")[:16])
    # sample_table = sample_page.extract_table()
    #print(sample_table)
    # df_table = pd.DataFrame(sample_table)
    # print(df_table.to_string(index=False, header=False))
    # df_table.to_csv("sample_table.csv", index=False, header=False)

