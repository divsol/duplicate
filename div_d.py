#!/usr/bin/env python
# coding: utf-8

# In[1]:


import streamlit as st
import pandas as pd
import os
import json
import io
import zipfile
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
    export_path = os.path.join(os.getcwd(), "access_export")
    os.makedirs(export_path, exist_ok=True)
    table_csvs = {}
    for table in tables:
        df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
        df.to_csv(os.path.join(export_path, f"{table}.csv"), index=False)
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

# UI Setup
st.set_page_config(page_title="Invoice Duplicate Checker", layout="centered")
st.title("Invoice Duplicate Checker")

last_db_path = load_config()
use_existing = False

if last_db_path and os.path.exists(last_db_path):
    st.info(f"Last used database: `{last_db_path}`\nLast updated: {get_db_modified_time(last_db_path)}")
    use_existing = st.radio("Choose database:", ["Use last", "Select new"]) == "Use last"

# Upload Access DB
if use_existing:
    db_path = last_db_path
else:
    db_file = st.file_uploader("Upload Access DB (.accdb)", type=["accdb"])
    if db_file:
        db_path = os.path.join(os.getcwd(), "current_db.accdb")
        with open(db_path, "wb") as f:
            f.write(db_file.getbuffer())
        save_config(db_path)
    else:
        st.stop()

# Convert Access to CSV and Load First Table
try:
    tables, table_data = convert_access_to_csv(db_path)
    selected_table = tables[0]
    master_df = table_data[selected_table][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
    master_df = generate_keys(master_df)
except Exception as e:
    st.error(f"Conversion error: {e}")
    st.stop()

# Upload Excel File
excel_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx", "xls"])
if not excel_file:
    st.stop()

# Load Excel and Match
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
    st.stop()

# Show Results
st.success("Duplicate check completed.")
st.dataframe(df)

# Download Results
excel_buffer = io.BytesIO()
final_df = raw_df.copy()
final_df['Duplicate'] = df['Duplicate']
final_df['Match Logic'] = df['Match Logic']
final_df.to_excel(excel_buffer, index=False)
st.download_button("📥 Download Results as Excel", data=excel_buffer.getvalue(), file_name="duplicates_report.xlsx")

# Merge non-duplicates into CSV (optional: show how to merge logic)
if "No" in df['Duplicate'].values:
    if st.button("🧩 Show UNIQUE Records"):
        unique_df = df[df['Duplicate'] == "No"][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
        st.write("These rows are not found in the Access database:")
        st.dataframe(unique_df)
else:
    st.warning("No unique records found.")



# In[ ]:




