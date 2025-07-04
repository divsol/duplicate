#!/usr/bin/env python
# coding: utf-8

# In[1]:


import streamlit as st
import pandas as pd
import os
import json
import io
import zipfile
import tempfile
from datetime import datetime
import pyodbc

CONFIG_FILE = "config.json"

def save_config(db_path):
    with open(CONFIG_FILE, 'w') as f:
        json.dump({"db_path": db_path}, f)

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f).get("db_path", "")
    return ""

def get_db_modified_time(path):
    return datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')

def generate_keys(df):
    df['key1'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key2'] = df['Invoice Number'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key3'] = df['Invoice Number'].astype(str) + '_' + df['Invoice Date'].astype(str) + '_' + df['Supplier Number'].astype(str)
    df['key4'] = df['Invoice Date'].astype(str) + '_' + df['Gross Amount'].astype(str) + '_' + df['Supplier Number'].astype(str) + '_' + df['Invoice Number'].astype(str)
    return df

def convert_access_to_csv(path):
    conn = pyodbc.connect(rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};')
    cursor = conn.cursor()
    tables = [t.table_name for t in cursor.tables(tableType='TABLE') if not t.table_name.startswith("MSys")]
    table_csvs = {}
    for table in tables:
        df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
        table_csvs[table] = df
    conn.close()
    return tables, table_csvs

def check_match(row, master_df):
    if row['key4'] in master_df['key4'].values:
        return pd.Series(['Yes', 'Date+Amount+Supplier+Number'])
    elif row['key1'] in master_df['key1'].values:
        return pd.Series(['Yes', 'Date+Amount+Supplier'])
    elif row['key2'] in master_df['key2'].values:
        return pd.Series(['Yes', 'Number+Amount+Supplier'])
    elif row['key3'] in master_df['key3'].values:
        return pd.Series(['Yes', 'Number+Date+Supplier'])
    else:
        return pd.Series(['No', 'UNIQUE'])

st.set_page_config(page_title="Invoice Duplicate Checker", layout="centered")
st.title("Invoice Duplicate Checker")

# Step 1: Upload Access DB
st.header("Step 1: Upload Access Database (.accdb)")
db_file = st.file_uploader("Upload Access DB (.accdb)", type=["accdb"])
if db_file is not None:
    # Save uploaded file to a temporary file
    temp_db_file = tempfile.NamedTemporaryFile(delete=False, suffix=".accdb")
    temp_db_file.write(db_file.getbuffer())
    temp_db_file.close()
    db_path = temp_db_file.name
    st.success(f"Database uploaded and saved to {db_path}")
else:
    st.info("Please upload a Microsoft Access database file (.accdb) above.")
    st.stop()

# Step 2: Load/Select Table
try:
    tables, table_data = convert_access_to_csv(db_path)
    if len(tables) == 0:
        st.error("No tables found in the Access database.")
        os.unlink(db_path)
        st.stop()
    selected_table = st.selectbox("Select table to use for duplicate checking:", tables)
    master_df = table_data[selected_table][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
    master_df = generate_keys(master_df)
except Exception as e:
    st.error(f"Error reading Access DB: {e}")
    os.unlink(db_path)
    st.stop()

# Step 3: Upload Excel File
st.header("Step 2: Upload Invoice Excel File")
excel_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx", "xls"])
if not excel_file:
    st.info("Please upload an Excel file containing invoices to be checked.")
    os.unlink(db_path)
    st.stop()

# Step 4: Check for Duplicates
try:
    raw_df = pd.read_excel(excel_file)
    raw_df.columns = raw_df.columns.str.strip()
    df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    df = df.dropna(subset=['Invoice Date'])
    df = generate_keys(df)
    df[['Duplicate', 'Match Logic']] = df.apply(lambda row: check_match(row, master_df), axis=1)
except Exception as e:
    st.error(f"Excel load error: {e}")
    os.unlink(db_path)
    st.stop()

# Step 5: Show Results
st.success("Duplicate check completed.")
st.dataframe(df)

# Step 6: Download Results
excel_buffer = io.BytesIO()
final_df = raw_df.copy()
final_df['Duplicate'] = df['Duplicate']
final_df['Match Logic'] = df['Match Logic']
final_df.to_excel(excel_buffer, index=False)
st.download_button("📥 Download Results as Excel", data=excel_buffer.getvalue(), file_name="duplicates_report.xlsx")

# Step 7: Show unique records
if "No" in df['Duplicate'].values:
    if st.button("🧩 Show UNIQUE Records"):
        unique_df = df[df['Duplicate'] == "No"][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
        st.write("These rows are not found in the Access database:")
        st.dataframe(unique_df)
else:
    st.warning("No unique records found.")

# Cleanup temp file
try:
    os.unlink(db_path)
except Exception:
    pass

# In[ ]:




