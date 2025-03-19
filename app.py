import streamlit as st
import pandas as pd
import openpyxl
from fuzzywuzzy import process  # For AI-based column mapping

# Standard column structure
STANDARD_COLUMNS = ["ISBN", "TITLE", "AUTHOR", "PUBLISHER", "STOCK", "CURRENCY", "PRICE", "COMPANY", "HANDLING"]

# Company details and handling conditions
COMPANY_OPTIONS = {
    "Adarsh": "2",
    "Adhya": "2",
    "UDH": "2",
    "RUPA": "2",
    "Prakash Delhi": "2",
    "Prakash Noida": "4",
    "IBD": "2",
    "GBD": "2",
    "ECP": "2",
    "VCP": "2"
}
COMPANY_ORDER = ["Adarsh", "GBD", "ECP", "IBD", "Prakash Delhi", "VCP", "Prakash Noida"]

# Stock filtering conditions
STOCK_CONDITIONS = {
    "Adarsh": 10,
    "Adhya": 2,
    "UDH": 0,
    "RUPA": 5,
    "Prakash Delhi": 4,
    "Prakash Noida": 4,
    "IBD": 3,
    "GBD": 5,
    "ECP": 0,
    "VCP": 3
}

# Hardcoded Currency Conversion Rates
CURRENCY_CONVERSION = {
    "USD": 90.60,
    "GBP": 113.80,
    "EUR": 94.70,
    "INR": 1,
    "RS": 1,
    "RS.": 1,
    "UKP": 113.80,
    "EU": 94.70,
    "€": 94.70,
    "£": 113.80,
    "$": 90.60
}

# AI-based column mapping function
def ai_column_mapping(columns, target_column):
    match, score = process.extractOne(target_column, columns)
    return match if score > 70 else None

# Streamlit interface
st.title("Excel File Processor 📊 (with AI Mapping, Stock Filters & Fixed Currency Rates)")
uploaded_files = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)

file_data = []
duplicate_assignment_errors = {}

if uploaded_files:
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, engine="openpyxl", dtype={"ISBN": str})
        except Exception:
            df = pd.read_excel(file, engine="openpyxl")
            if "ISBN" in df.columns:
                df["ISBN"] = df["ISBN"].astype(str)

        df.columns = df.columns.astype(str)
        original_columns = df.columns.tolist()

        # AI-based and manual column mapping
        mapping_suggestion = {}
        for col in original_columns:
            assigned = False
            mapped_col = ai_column_mapping(STANDARD_COLUMNS, col)
            if mapped_col:
                mapping_suggestion[col] = mapped_col
                assigned = True
            if not assigned:
                mapping_suggestion[col] = None

        # Display column mapping
        st.subheader(f"Column Mapping for {file.name}")
        final_mapping = {}
        assigned_standard = []
        duplicate_assigned = []

        for col in original_columns:
            selected = st.selectbox(
                f"Assign '{col}' to:",
                ["(Leave Blank)"] + STANDARD_COLUMNS,
                index=(STANDARD_COLUMNS.index(mapping_suggestion[col]) + 1) if mapping_suggestion[col] in STANDARD_COLUMNS else 0,
                key=f"{file.name}_mapping_{col}"
            )
            final_mapping[col] = selected
            if selected != "(Leave Blank)":
                if selected in assigned_standard:
                    duplicate_assigned.append(selected)
                else:
                    assigned_standard.append(selected)

        # Error handling for duplicates
        if duplicate_assigned:
            duplicate_assignment_errors[file.name] = duplicate_assigned
            st.error(f"Duplicate assignment in {file.name}: {', '.join(duplicate_assigned)}. Please fix them.")

        company_selection = st.selectbox(f"Select COMPANY for {file.name}:", list(COMPANY_OPTIONS.keys()), key=f"{file.name}_company")
        file_data.append((df, final_mapping, company_selection, file.name))

    # Prevent merging until errors are resolved
    if duplicate_assignment_errors:
        st.error("Please resolve all duplicate column assignments before merging.")
    else:
        if st.button("Merge Files 🚀"):
            merged_dfs = []

            for df, mapping, company, fname in file_data:
                df = df.copy()
                df.rename(columns=mapping, inplace=True)

                selected_cols = [v for v in mapping.values() if v != "(Leave Blank)"]
                if not selected_cols:
                    st.warning(f"No columns selected for {fname}; skipping this file.")
                    continue

                df = df[selected_cols]

                # Stock filtering
                if "STOCK" in df.columns and company in STOCK_CONDITIONS:
                    min_stock = STOCK_CONDITIONS[company]
                    df = df[df["STOCK"] >= min_stock]

                df["COMPANY"] = company
                df["HANDLING"] = COMPANY_OPTIONS[company]
                df.reset_index(drop=True, inplace=True)
                merged_dfs.append(df)

            if merged_dfs:
                merged_df = pd.concat(merged_dfs, ignore_index=True)
                merged_df = merged_df.reindex(columns=[col for col in STANDARD_COLUMNS if col in merged_df.columns])

                # Data cleaning and validation
                # Clean and standardize ISBN
            if "ISBN" in merged_df.columns:
                merged_df["ISBN"] = (
                    merged_df["ISBN"]
                    .astype(str)                        # Ensure ISBN is treated as text
                    .str.strip()                        # Remove leading/trailing spaces
                    .str.replace(r'[^\dXx]', '', regex=True)  # Remove non-numeric/letter characters
                    .replace(["nan", "NaN", "", "0"], pd.NA)  # Treat invalid entries as null
                )

                # Drop duplicate ISBN entries
                merged_df.drop_duplicates(subset=["ISBN"], keep="first", inplace=True)


                for col in ["PRICE", "STOCK"]:
                    if col in merged_df.columns:
                        merged_df = merged_df[merged_df[col] != 0]

                # Currency conversion
                if "CURRENCY" in merged_df.columns and "PRICE" in merged_df.columns:
                    merged_df["CURRENCY"] = merged_df["CURRENCY"].astype(str).str.strip().str.upper()
                    merged_df["PRICE"] = merged_df.apply(
                        lambda row: row["PRICE"] * CURRENCY_CONVERSION.get(row["CURRENCY"], 1.0),
                        axis=1
                    )
                    merged_df["CURRENCY"] = "INR"

                # Sorting by company order
                merged_df["COMPANY"] = pd.Categorical(merged_df["COMPANY"], categories=COMPANY_ORDER, ordered=True)
                merged_df.sort_values(by="COMPANY", inplace=True)

                # Display and download processed data
                st.subheader("Final Processed Data")
                st.dataframe(merged_df)
                csv = merged_df.to_csv(index=False).encode("utf-8")
                st.download_button("📥 Download Processed File", csv, "processed_data.csv", "text/csv")
            else:
                st.error("No valid data to merge.")
