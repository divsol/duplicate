#!/usr/bin/env python
# coding: utf-8

# In[1]:


import streamlit as st
import pandas as pd
import pyodbc
import os
import json
import io
from datetime import datetime

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

def connect_access_db(path):
    return pyodbc.connect(rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};')

def check_match(row):
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

# Streamlit UI setup
st.set_page_config(page_title="Invoice Duplicate Checker", layout="centered")
st.title("Invoice Duplicate Checker")

# Load previous DB
last_db_path = load_config()
use_existing = False

if last_db_path and os.path.exists(last_db_path):
    st.info(f"Last used database: `{last_db_path}`\nLast updated: {get_db_modified_time(last_db_path)}")
    use_existing = st.radio("Choose database:", ["Use last", "Select new"]) == "Use last"

if use_existing:
    db_path = last_db_path
else:
    db_file = st.file_uploader("Upload MS Access DB (.accdb)", type=["accdb"])
    if db_file:
        db_path = os.path.join(os.getcwd(), "current_db.accdb")
        with open(db_path, "wb") as f:
            f.write(db_file.getbuffer())
        save_config(db_path)
    else:
        st.stop()

# Upload Excel file
excel_file = st.file_uploader("Upload Invoice Excel File", type=["xlsx", "xls"])
if not excel_file:
    st.stop()

# Load Access DB
try:
    conn = connect_access_db(db_path)
    cursor = conn.cursor()
    tables = [t.table_name for t in cursor.tables(tableType='TABLE') if not t.table_name.startswith("MSys")]
    table = tables[0]
    query = f"SELECT [Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number] FROM [{table}]"
    master_df = pd.read_sql(query, conn)
    master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
    master_df = generate_keys(master_df)
except Exception as e:
    st.error(f"Access DB error: {e}")
    st.stop()

# Load Excel input
try:
    raw_df = pd.read_excel(excel_file)
    raw_df.columns = raw_df.columns.str.strip()
    df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']].copy()
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    df = df.dropna(subset=['Invoice Date'])
    df = generate_keys(df)
except Exception as e:
    st.error(f"Excel load error: {e}")
    st.stop()

# Matching Logic
df[['Duplicate', 'Match Logic']] = df.apply(check_match, axis=1)
st.success("Duplicate check completed.")
st.dataframe(df)

# Export to Excel with original columns preserved
excel_buffer = io.BytesIO()
final_df = raw_df.copy()
final_df['Duplicate'] = df['Duplicate']
final_df['Match Logic'] = df['Match Logic']
final_df.to_excel(excel_buffer, index=False)

st.download_button(
    "📥 Download Results as Excel",
    data=excel_buffer.getvalue(),
    file_name="duplicates_report.xlsx"
)

# Merge non-duplicates
if "No" in df['Duplicate'].values:
    if st.button("🧩 Merge UNIQUE Records into Access DB"):
        unique_df = df[df['Duplicate'] == "No"][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
        try:
            for _, row in unique_df.iterrows():
                cursor.execute(
                    f"""INSERT INTO [{table}] ([Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number])
                        VALUES (?, ?, ?, ?)""",
                    row['Invoice Number'], row['Invoice Date'], row['Gross Amount'], row['Supplier Number']
                )
            conn.commit()
            st.success("Unique rows merged into Access database.")
        except Exception as e:
            st.error(f"Merge error: {e}")
else:
    st.warning("No unique records to merge.")

conn.close()



# In[ ]:




