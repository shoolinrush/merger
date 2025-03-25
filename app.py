import streamlit as st
import pandas as pd
import openpyxl
from fuzzywuzzy import process  # For AI-based column mapping
import time

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

st.title("Excel File Processor 📊 (with AI Mapping, Stock Filters & Fixed Currency Rates)")

# Upload multiple Excel files
uploaded_files = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)

# Lists for storing file related data and error information
file_data = []
duplicate_assignment_errors = {}

if uploaded_files:
    for file in uploaded_files:
        try:
            # Read file; ensure ISBN is read as string if available
            df = pd.read_excel(file, engine="openpyxl", dtype={"ISBN": str})
        except Exception:
            df = pd.read_excel(file, engine="openpyxl")
            if "ISBN" in df.columns:
                df["ISBN"] = df["ISBN"].astype(str)
        
        df.columns = df.columns.astype(str)
        original_columns = df.columns.tolist()

        # AI-based suggestion for column mapping
        mapping_suggestion = {}
        for col in original_columns:
            mapped_col = ai_column_mapping(STANDARD_COLUMNS, col)
            mapping_suggestion[col] = mapped_col if mapped_col in STANDARD_COLUMNS else None

        st.subheader(f"Column Mapping for {file.name}")
        final_mapping = {}
        assigned_standard = []
        duplicate_assigned = []

        # Create select boxes for manual override of the AI suggestion
        for col in original_columns:
            default_index = STANDARD_COLUMNS.index(mapping_suggestion[col]) + 1 if mapping_suggestion[col] in STANDARD_COLUMNS else 0
            selected = st.selectbox(
                f"Assign '{col}' to:",
                ["(Leave Blank)"] + STANDARD_COLUMNS,
                index=default_index,
                key=f"{file.name}_mapping_{col}"
            )
            final_mapping[col] = selected
            # Check for duplicates (if a standard column has been already assigned)
            if selected != "(Leave Blank)":
                if selected in assigned_standard:
                    duplicate_assigned.append(selected)
                else:
                    assigned_standard.append(selected)

        if duplicate_assigned:
            duplicate_assignment_errors[file.name] = duplicate_assigned
            st.error(f"Duplicate assignment in {file.name}: {', '.join(duplicate_assigned)}. Please fix them.")

        # Select company for this file
        company_selection = st.selectbox(
            f"Select COMPANY for {file.name}:",
            list(COMPANY_OPTIONS.keys()),
            key=f"{file.name}_company"
        )
        file_data.append((df, final_mapping, company_selection, file.name))

    if duplicate_assignment_errors:
        st.error("Please resolve all duplicate column assignments before merging.")
    else:
        if st.button("Merge Files 🚀"):
            merged_dfs = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            total_files = len(file_data)
            for idx, (df, mapping, company, fname) in enumerate(file_data):
                status_text.text(f"Processing {fname}...")
                df = df.copy()
                df.rename(columns=mapping, inplace=True)

                # Filter out columns not selected
                selected_cols = [v for v in mapping.values() if v != "(Leave Blank)"]
                if not selected_cols:
                    st.warning(f"No columns selected for {fname}; skipping this file.")
                    continue
                df = df[selected_cols]

                # Apply stock filter if applicable
                if "STOCK" in df.columns and company in STOCK_CONDITIONS:
                    try:
                        # Ensure STOCK is numeric
                        df["STOCK"] = pd.to_numeric(df["STOCK"], errors="coerce")
                        min_stock = STOCK_CONDITIONS[company]
                        df = df[df["STOCK"] >= min_stock]
                    except Exception as e:
                        st.error(f"Error processing STOCK for {fname}: {e}")
                        continue

                # Add company and handling columns
                df["COMPANY"] = company
                df["HANDLING"] = COMPANY_OPTIONS.get(company, "0")
                df.reset_index(drop=True, inplace=True)
                merged_dfs.append(df)

                # Update progress
                progress_bar.progress((idx + 1) / total_files)
                time.sleep(0.2)  # simulate processing time

            if merged_dfs:
                with st.spinner("Merging and cleaning data..."):
                    merged_df = pd.concat(merged_dfs, ignore_index=True)

                    # Reindex the merged dataframe to include only the standard columns available
                    merged_df = merged_df.reindex(columns=[col for col in STANDARD_COLUMNS if col in merged_df.columns])

                    # Data cleaning for ISBN: stripping, removing unwanted characters, and marking invalids as null
                    if "ISBN" in merged_df.columns:
                        merged_df["ISBN"] = (
                            merged_df["ISBN"]
                            .astype(str)
                            .str.strip()
                            .str.replace(r'[^\dXx]', '', regex=True)
                            .replace(["nan", "NaN", "", "0"], pd.NA)
                        )
                        # Drop duplicate ISBN entries
                        merged_df.drop_duplicates(subset=["ISBN"], keep="first", inplace=True)

                    # Remove rows where PRICE or STOCK are 0 (if they exist)
                    for col in ["PRICE", "STOCK"]:
                        if col in merged_df.columns:
                            merged_df = merged_df[merged_df[col] != 0]

                    # Currency conversion: multiply PRICE by fixed conversion rate and set currency to INR
                    if "CURRENCY" in merged_df.columns and "PRICE" in merged_df.columns:
                        merged_df["CURRENCY"] = merged_df["CURRENCY"].astype(str).str.strip().str.upper()
                        merged_df["PRICE"] = merged_df.apply(
                            lambda row: row["PRICE"] * CURRENCY_CONVERSION.get(row["CURRENCY"], 1.0),
                            axis=1
                        )
                        merged_df["CURRENCY"] = "INR"

                    # Sort by company order (using a categorical order)
                    merged_df["COMPANY"] = pd.Categorical(merged_df["COMPANY"], categories=COMPANY_ORDER, ordered=True)
                    merged_df.sort_values(by="COMPANY", inplace=True)
                    merged_df.reset_index(drop=True, inplace=True)
                    time.sleep(0.5)
                
                st.success("Files merged successfully!")
                st.subheader("Final Processed Data")
                st.dataframe(merged_df)
                csv = merged_df.to_csv(index=False).encode("utf-8")
                st.download_button("📥 Download Processed File", csv, "processed_data.csv", "text/csv")
            else:
                st.error("No valid data to merge.")
