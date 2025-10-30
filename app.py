import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from io import BytesIO
import requests
from openpyxl.utils import range_boundaries
import xlsxwriter

# Page config and heading (exact text requested)
st.set_page_config(page_title="Personal Timetable Creator- Trimester 5, 2nd Half", layout="centered")
st.title("Personal Timetable Creator- Trimester 5, 2nd Half")

# The exact instruction text required
st.write("""
""" + " Step 1: Select your subjects from the list\nStep 2: Click on generate timetable button\nStep 3: Click on Download button to download the excel file" + """
""")

# --- Download and preprocess the timetable sheet (run once on load) ---
# Google Drive File Link (URL) of timetable (same URL you provided)
url = "https://docs.google.com/spreadsheets/d/1hxMVAdZM-aaHY1IDy7Hg8wLPdevSEhVx/edit?usp=sharing&ouid=106900160560444308561&rtpof=true&sd=true"

# Extract the File Id and download xlsx
file_id = url.split('/')[-2]
download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
response = requests.get(download_url)
wb = load_workbook(BytesIO(response.content))
ws = wb.active

# Handle merged cells - Unmerge and fill values (preserve your logic)
merged_info = []
for merged_range in list(ws.merged_cells.ranges):
    min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
    value = ws.cell(row=min_row, column=min_col).value
    merged_info.append((min_row, min_col, max_row, max_col, value))

for merged_range in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(merged_range))

for min_row, min_col, max_row, max_col, value in merged_info:
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).value = value

# Load into DataFrame
data = list(ws.values)
tt = pd.DataFrame(data)

# Set headers
tt.columns = tt.iloc[0]
tt = tt[1:].reset_index(drop=True)

# Remove columns with same name and same values (preserve your function)
def drop_duplicate_named_columns(df):
    keep = []
    seen = set()
    for col in df.columns:
        col_data = tuple(df[col])
        col_key = (col, col_data)
        if col_key not in seen:
            seen.add(col_key)
            keep.append(col)
    return df[keep]

tt = drop_duplicate_named_columns(tt)
tt = tt.loc[:, ~tt.T.duplicated(keep='first')]

# Rename first column to 'Day/time' if needed
if tt.columns[0] != 'Day/time':
    tt.rename(columns={tt.columns[0]: 'Day/time'}, inplace=True)

# Remove duplicate (column name, row header, value) by making later ones NaN
row_header_col_index = 0  # 'Day/time' is the first column
row_headers = tt.iloc[:, row_header_col_index]

seen = set()
for col_index in range(1, tt.shape[1]):
    for row_index in range(tt.shape[0]):
        row_header_value = row_headers.iat[row_index]
        cell_value = tt.iat[row_index, col_index]
        col_name = tt.columns[col_index]
        key = (col_name, row_header_value, cell_value)
        if key in seen:
            tt.iat[row_index, col_index] = np.nan
        else:
            seen.add(key)

# Replace string 'None' and Python None with actual NaN
tt.replace(to_replace=["None", None], value=np.nan, inplace=True)

# Drop columns with all null values
tt.dropna(axis=1, how='all', inplace=True)

# Ensure all column names are strings
tt.columns = tt.columns.astype(str)

# Rename duplicate column names to 'Unnamed 1', 'Unnamed 2', ...
new_cols = []
counter = 1
seen_cols = set()
for col in tt.columns:
    if col in seen_cols:
        new_cols.append(f"Unnamed {counter}")
        counter += 1
    else:
        seen_cols.add(col)
        new_cols.append(col)
tt.columns = new_cols

# Identify 'Unnamed' columns to concatenate
unnamed_cols = [col for col in tt.columns if col.startswith('Unnamed')]

# Iterate and concatenate 'Unnamed' columns to their left columns
for col in unnamed_cols:
    left_col_index = tt.columns.get_loc(col) - 1
    left_col = tt.columns[left_col_index]
    for index, row in tt.iterrows():
        if pd.notna(row[col]):
            if pd.notna(row[left_col]):
                tt.loc[index, left_col] = f"{row[left_col]} {row[col]}"
            else:
                tt.loc[index, left_col] = row[col]

# Remove the 'Unnamed' columns
tt = tt.drop(columns=unnamed_cols)

# -------------------------
# DYNAMIC SUBJECT LIST EXTRACTION
# -------------------------
# Your cells look like: "ES-1 SM CR-17" where first token is subject (ES-1).
# We'll extract the first token from each non-null cell (excluding 'Day/time') and collect unique values.

subject_set = set()
for col in tt.columns:
    if col == 'Day/time':
        continue
    for val in tt[col].dropna():
        # keep the exact tokenization you described: first token separated by whitespace
        tokens = str(val).split()
        if len(tokens) >= 1:
            subject_set.add(tokens[0])

# Sort for nicer display
subject_list = sorted(subject_set)

# If no subjects found, show warning and still allow manual input
if not subject_list:
    st.warning("No subject tokens found automatically in the timetable. You can input subject names manually.")
    my_subjects = st.text_input("Enter subjects (comma-separated):")
    if my_subjects:
        my_subjects = [s.strip() for s in my_subjects.split(",") if s.strip()]
    else:
        my_subjects = []
else:
    my_subjects = st.multiselect("Select your subjects", options=subject_list)

# Button to generate timetable (preserve all matching logic)
if st.button("Generate Timetable"):
    if not my_subjects:
        st.warning("Please select at least one subject to generate your timetable.")
    else:
        # Create an empty DataFrame for personal timetable with same columns as tt
        personal_tt = pd.DataFrame(columns=tt.columns)

        # Matching block: EXACTLY as you provided (no logic change)
        for index, row in tt.iterrows():
            for col in tt.columns:
                if col != 'Day/time':
                    cell_value = row[col]
                    if pd.notna(cell_value):
                        cell_subjects = str(cell_value).split()  # Exact match using space split
                        for subject in my_subjects:
                            if subject in cell_subjects:
                                personal_tt.loc[index, col] = cell_value
                                personal_tt.loc[index, 'Day/time'] = row['Day/time']
                                break

        # Ensure all unique 'Day/time' entries from the original tt are in personal_tt
        unique_days = tt[['Day/time']].drop_duplicates()
        personal_tt = pd.merge(unique_days, personal_tt, on='Day/time', how='left')

        # Merge rows with the same 'Day/time' and aggregate non-null values
        personal_tt = personal_tt.groupby('Day/time').agg(lambda x: ' '.join(x.dropna())).reset_index()

        # Reorder personal_tt to match the original date order
        original_day_order = tt['Day/time'].drop_duplicates()
        personal_tt['Day/time'] = pd.Categorical(personal_tt['Day/time'], categories=original_day_order, ordered=True)
        personal_tt = personal_tt.sort_values('Day/time').reset_index(drop=True)

        # Display the timetable in Streamlit
        st.success("✅ Personal timetable generated successfully!")
        st.dataframe(personal_tt)

        # Save the personal timetable to an Excel file in memory with same formatting logic
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            personal_tt.to_excel(writer, sheet_name='Personal Timetable', index=False)

            workbook  = writer.book
            worksheet = writer.sheets['Personal Timetable']

            # Set default format for all cells
            cell_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Set bold format
            bold_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Apply format to all cells, applying bold format to header and Day/time column
            for row_num in range(len(personal_tt) + 1): # +1 for header row
                for col_num in range(len(personal_tt.columns)):
                    cell_value = personal_tt.iloc[row_num-1, col_num] if row_num > 0 else personal_tt.columns[col_num]
                    if row_num == 0 or col_num == 0: # Apply bold format to header row and first column
                        worksheet.write(row_num, col_num, cell_value, bold_format)
                    else:
                        worksheet.write(row_num, col_num, cell_value, cell_format)

            # Set row height
            for row_num in range(len(personal_tt) + 1): # +1 for header row
                worksheet.set_row(row_num, 32.4)

            # Set column width
            for col_num in range(len(personal_tt.columns)):
                worksheet.set_column(col_num, col_num, 20)

        output.seek(0)

        # Download button
        st.download_button(
            label="📥 Download Timetable",
            data=output,
            file_name="personal_timetable_formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
