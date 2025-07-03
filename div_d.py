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
    return df

def connect_access_db(path):
    return pyodbc.connect(rf'Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={path};')

st.set_page_config(page_title=" Invoice Duplicate Checker", layout="centered")
st.title("Invoice Duplicate Checker")

# Load config
last_db_path = load_config()
use_existing = False

if last_db_path and os.path.exists(last_db_path):
    st.info(f"Last used database: `{last_db_path}`\n\n Last updated: {get_db_modified_time(last_db_path)}")
    use_existing = st.radio("Use the last saved database or select a new one:", ["Use last", "Select new"]) == "Use last"

if use_existing:
    db_path = last_db_path
else:
    db_file = st.file_uploader("Upload MS Access Database (.accdb)", type=["accdb"])
    if db_file:
        db_path = os.path.join(os.getcwd(), "current_db.accdb")
        with open(db_path, "wb") as f:
            f.write(db_file.getbuffer())
        save_config(db_path)
    else:
        st.stop()

# Excel upload
excel_file = st.file_uploader("Upload Excel Invoice File", type=["xlsx", "xls"])
if not excel_file:
    st.stop()


try:
    conn = connect_access_db(db_path)
    cursor = conn.cursor()
    tables = [t.table_name for t in cursor.tables(tableType='TABLE') if not t.table_name.startswith("MSys")]
    table = tables[0]

    db_query = f"SELECT [Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number] FROM [{table}]"
    master_df = pd.read_sql(db_query, conn)
    master_df['Invoice Date'] = pd.to_datetime(master_df['Invoice Date'], errors='coerce')
except Exception as e:
    st.error(f"Error loading Access DB: {e}")
    st.stop()

try:
    raw_df = pd.read_excel(excel_file)
    raw_df.columns = raw_df.columns.str.strip()
    df = raw_df[['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    df = df.dropna(subset=['Invoice Date'])
except Exception as e:
    st.error(f"Error loading Excel: {e}")
    st.stop()

df = generate_keys(df)
master_df = generate_keys(master_df)

def check_match(row):
    if row['key1'] in master_df['key1'].values:
        return 'Date+Amount+Supplier'
    elif row['key2'] in master_df['key2'].values:
        return 'Number+Amount+Supplier'
    elif row['key3'] in master_df['key3'].values:
        return 'Number+Date+Supplier'
    else:
        return 'UNIQUE'

df['Status'] = df.apply(check_match, axis=1)
st.success(" Duplicate check completed.")
st.dataframe(df)


excel_buffer = io.BytesIO()
df.to_excel(excel_buffer, index=False)
st.download_button(" Download Report as Excel", data=excel_buffer.getvalue(), file_name="duplicates_report.xlsx")


if "UNIQUE" in df['Status'].values:
    if st.button(" Merge UNIQUE rows into Access Database"):
        unique_df = df[df['Status'] == "UNIQUE"][['Invoice Number', 'Invoice Date', 'Gross Amount', 'Supplier Number']]
        try:
            for _, row in unique_df.iterrows():
                cursor.execute(
                    f"""INSERT INTO [{table}] ([Invoice Number], [Invoice Date], [Gross Amount], [Supplier Number])
                        VALUES (?, ?, ?, ?)""",
                    row['Invoice Number'], row['Invoice Date'], row['Gross Amount'], row['Supplier Number']
                )
            conn.commit()
            st.success(" Unique rows merged into the Access database.")
        except Exception as e:
            st.error(f"Merge error: {e}")
else:
    st.warning("No unique records to merge.")

conn.close()


# In[ ]:




