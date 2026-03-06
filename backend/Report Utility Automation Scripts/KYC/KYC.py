import pandas as pd

def KYC_main(excel_file, output_folder_path):
    # Define file paths
    output_file = "KYC_Result.xlsx"
    full_path = rf"{output_folder_path}\{output_file}"

    # Load Sheet2, Sheet3, and Sheet4
    df_sheet2 = pd.read_excel(excel_file, sheet_name="Sheet2", engine="openpyxl")
    df_sheet3 = pd.read_excel(excel_file, sheet_name="Sheet3", engine="openpyxl")
    df_sheet4 = pd.read_excel(excel_file, sheet_name="Sheet4", engine="openpyxl")

    # KPI 1: Same LCIN, different PAN No
    kpi1_df = df_sheet2.groupby("LCIN").filter(lambda x: x["PAN No"].nunique() > 1)

    # KPI 2: Same LCIN, different Final Risk Rating
    kpi2_df = df_sheet2.groupby("LCIN").filter(lambda x: x["Final Risk Rating"].nunique() > 1)

    # KPI 3: Map PAN Type using 4th character of PAN No
    mapping_dict = {}
    for _, row in df_sheet3.iterrows():
        parts = row["PAN 4 character list"].split("–")
        if len(parts) == 2:
            key = parts[0].strip()
            value = parts[1].strip()
            mapping_dict[key] = value

    df_sheet2["PAN Type"] = df_sheet2["PAN No"].astype(str).str[3].map(mapping_dict)

    # KPI 4: Check if any word in DBSIC Description matches any word in any Description from Sheet4
    keywords = set()
    for desc in df_sheet4["Description"].dropna():
        for word in desc.replace(",", " ").split():
            keywords.add(word.strip().lower())

    def check_high_risk(description):
        if pd.isna(description):
            return ""
        words = description.replace(",", " ").split()
        for word in words:
            if word.strip().lower() in keywords:
                return "High Risk"
        return ""

    df_sheet2["Risk Rating based on Desc"] = df_sheet2["DBSIC Description"].apply(check_high_risk)

    # Save all KPI results into one Excel file with separate sheets
    with pd.ExcelWriter(full_path, engine="openpyxl") as writer:
        kpi1_df.to_excel(writer, sheet_name="KPI_1", index=False)
        kpi2_df.to_excel(writer, sheet_name="KPI_2", index=False)
        df_sheet2.to_excel(writer, sheet_name="KPI_3", index=False)
        df_sheet2.to_excel(writer, sheet_name="KPI_4", index=False)

    return full_path
