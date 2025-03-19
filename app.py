import streamlit as st
import pandas as pd
import openpyxl
import requests

STANDARD_COLUMNS = ["ISBN", "TITLE", "AUTHOR", "PUBLISHER", "STOCK", "CURRENCY", "PRICE", "COMPANY", "HANDLING"]
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

API_KEY = "de86638281a545de980629c5"
API_BASE_URL = "https://v6.exchangerate-api.com/v6"

def get_rate_to_inr(base_currency):
    code = base_currency.strip().upper()
    alias = {"RS": "INR", "RS.": "INR", "UKP": "GBP", "EU": "EUR", "€": "EUR", "£": "GBP", "$": "USD"}
    if code in alias:
        code = alias[code]
    if code == "INR":
        return 1.0
    url = f"{API_BASE_URL}/{API_KEY}/latest/{code}"
    try:
        response = requests.get(url)
        data = response.json()
        if data.get("result") == "success":
            rate = data["conversion_rates"].get("INR", None)
            return rate if rate is not None else 1.0
        else:
            st.warning(f"API error for base {base_currency}: {data.get('error-type', 'Unknown error')}. Using fallback rate 1.0")
            return 1.0
    except Exception as e:
        st.warning(f"Error fetching rate for {base_currency}: {e}. Using fallback rate 1.0")
        return 1.0

st.title("Excel File Processor 📊")
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
        mapping_suggestion = {}
        for col in original_columns:
            assigned = False
            if col.strip().lower() in ["bookname", "name"]:
                mapping_suggestion[col] = "TITLE"
                assigned = True
            if not assigned:
                for std in STANDARD_COLUMNS:
                    if std.lower() in col.lower() and std not in mapping_suggestion.values():
                        mapping_suggestion[col] = std
                        assigned = True
                        break
            if not assigned:
                mapping_suggestion[col] = None
        st.subheader(f"Column Mapping for {file.name}")
        final_mapping = {}
        assigned_standard = []
        duplicate_assigned = []
        for col in original_columns:
            selected = st.selectbox(f"Assign '{col}' to:", ["(Leave Blank)"] + STANDARD_COLUMNS, 
                                    index=(STANDARD_COLUMNS.index(mapping_suggestion[col]) + 1) if mapping_suggestion[col] in STANDARD_COLUMNS else 0, 
                                    key=f"{file.name}_mapping_{col}")
            final_mapping[col] = selected
            if selected != "(Leave Blank)":
                if selected in assigned_standard:
                    duplicate_assigned.append(selected)
                else:
                    assigned_standard.append(selected)
        if duplicate_assigned:
            duplicate_assignment_errors[file.name] = duplicate_assigned
            st.error(f"Duplicate assignment in {file.name}: {', '.join(duplicate_assigned)}. Please fix them.")
        company_selection = st.selectbox(f"Select COMPANY for {file.name}:", list(COMPANY_OPTIONS.keys()), key=f"{file.name}_company")
        file_data.append((df, final_mapping, company_selection, file.name))
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
                df["COMPANY"] = company
                df["HANDLING"] = COMPANY_OPTIONS[company]
                df.reset_index(drop=True, inplace=True)
                merged_dfs.append(df)
            if merged_dfs:
                merged_df = pd.concat(merged_dfs, ignore_index=True)
                merged_df = merged_df.reindex(columns=[col for col in STANDARD_COLUMNS if col in merged_df.columns])

                if "ISBN" in merged_df.columns:
                    merged_df.drop_duplicates(subset=["ISBN"], keep="first", inplace=True)
                    merged_df["ISBN"] = merged_df["ISBN"].astype(str).replace(["nan", "NaN"], pd.NA)

                for col in ["PRICE", "STOCK"]:
                    if col in merged_df.columns:
                        merged_df[col] = pd.to_numeric(merged_df[col], errors="coerce")

                for col in ["PRICE", "ISBN", "CURRENCY", "STOCK"]:
                    if col in merged_df.columns:
                        merged_df[col] = merged_df[col].replace("", pd.NA)
                merged_df.dropna(subset=["PRICE", "ISBN", "CURRENCY", "STOCK"], inplace=True)

                for col in ["PRICE", "STOCK"]:
                    if col in merged_df.columns:
                        merged_df = merged_df[merged_df[col] != 0]

                if "ISBN" in merged_df.columns:
                    merged_df = merged_df[merged_df["ISBN"].str.strip() != "0"]

                if "CURRENCY" in merged_df.columns:
                    merged_df["CURRENCY"] = merged_df["CURRENCY"].astype(str)
                    merged_df = merged_df[~merged_df["CURRENCY"].str.strip().str.match(r'^-?\d+(\.\d+)?$')]

                if "CURRENCY" in merged_df.columns and "PRICE" in merged_df.columns:
                    unique_currencies = merged_df["CURRENCY"].dropna().str.strip().str.upper().unique()
                    conversion_dict = {cur: get_rate_to_inr(cur) for cur in unique_currencies}
                    merged_df["PRICE"] = merged_df.apply(lambda row: row["PRICE"] * conversion_dict.get(row["CURRENCY"].strip().upper(), 1.0), axis=1)
                    merged_df["CURRENCY"] = "INR"

                merged_df["COMPANY"] = pd.Categorical(merged_df["COMPANY"], categories=COMPANY_ORDER, ordered=True)
                merged_df.sort_values(by="COMPANY", inplace=True)

                st.subheader("Final Processed Data")
                st.dataframe(merged_df)
                csv = merged_df.to_csv(index=False).encode("utf-8")
                st.download_button("📥 Download Processed File", csv, "processed_data.csv", "text/csv")
            else:
                st.error("No valid data to merge.")
