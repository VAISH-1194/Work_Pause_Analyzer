from flask import Flask, request, render_template, send_file
import pandas as pd
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)

@app.route('/')
def upload_form():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    
    if file:
        df_initial = pd.read_excel(file, header=None)
        # Your processing code here

        # Drop all rows that are completely empty
        df_non_empty = df_initial.dropna(how='all')

        # Find the first non-empty row
        first_non_empty_row = df_non_empty.index[0]

        # Read the Excel file again from the first non-empty row
        df = pd.read_excel(file, header=first_non_empty_row)

        # Display the first few rows of the DataFrame to verify
        df = pd.read_excel(file, skiprows=9)

        # Dropping the Unnamed Columns
        df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)

        # Add a copy column of Records in the name Punch Records
        df['Records'] = df['Punch Records']

# Function to change names and terms in records
    def change_name(records):
        if pd.isna(records):
            return records

        # Replace 'BD' with 'ED'
        entries = records.split(',')
        entries = [entry.replace('BD', 'ED') for entry in entries]

        # Replace 'Main Entrance' with 'ED' and 'Exit' with 'ED'
        entries = [entry.replace('Main Entrance', 'ED').replace('Exit', 'ED') for entry in entries]

        return ', '.join(entries)

    # Apply the function to 'Records' column
    df['Records'] = df['Records'].apply(change_name)

    # Function to filter punch records
    def filter_punch_records(record):
        if pd.isna(record):
            return record

        entries = record.split(',')
        valid_entries = [entry for entry in entries if ('in' in entry or 'out' in entry)]

        return ','.join(valid_entries)

    df['Records'] = df['Records'].apply(filter_punch_records)

    # Replace 'NaN' with pd.NA in 'Punch Records' and 'Records' columns
    df['Punch Records'].replace('NaN', pd.NA, inplace=True)
    df['Records'].replace('NaN', pd.NA, inplace=True)

    # Adding the "Status" column based on the "Records" column
    df['Employee Status'] = df['Records'].apply(lambda x: 'Present' if pd.notna(x) and x != '' else 'Absent')

    def update_status_based_on_records(records, punch_records):
        if pd.isna(records) or records.strip() == '':
            return 'Absent'

        entries = records.split(', ')

        if len(entries) == 1:
            return 'Punch records missing'

        if 'out' in entries[0]:
            return 'Punch records missing'

        if len(entries) % 2 != 0:
            return 'Punch records missing'

        # Check for consecutive 'in' or 'out' entries
        for i in range(1, len(entries)):
            if ('in' in entries[i] and 'in' in entries[i-1]) or \
            ('out' in entries[i] and 'out' in entries[i-1]):
                return 'Punch records missing'

        return 'Valid Records'

    # Function to mark columns empty based on specific words
    def mark_columns_empty(row):
        words_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                        'Punch Records', 'Emp Code:', 'Employee Name :', 'Department:', 'Records']
        
        if row.name > 0:  # Check if row index is greater than 0 (excluding headers)
            for word in words_to_check:
                if pd.notna(row[word]) and any(word in str(cell) for cell in row):
                    row['Employee Status'] = " "
                    row['Records Status'] = " "
                    row['Break Time'] = " "
                    return row
        return row

    # Apply the function to update the columns
    df = df.apply(mark_columns_empty, axis=1)

    # Apply the function to update 'Records Status' column
    df['Records Status'] = df.apply(lambda row: update_status_based_on_records(row['Records'], row['Punch Records']), axis=1)

    # Columns to check for NaN or words and alphabets
    cols_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                    'Punch Records', 'Emp Code:', 'Employee Name :', 'Department:', 'Records']

    def check_nan_and_update_status(row, cols_to_check):
        if all(pd.isna(row[col]) for col in cols_to_check):
            row['Employee Status'] = " "
            row['Records Status'] = " "
            row['Break Time'] = " "
        return row

    # Apply the function to update 'Employee Status' column
    df = df.apply(lambda row: check_nan_and_update_status(row, cols_to_check), axis=1)

    # Reorder columns to ensure "Employee Status", "Records Status", and "Break Time (minutes)" are last
    columns_order = [col for col in df.columns if col not in ['Employee Status', 'Records Status', 'Break Time']]
    columns_order += ['Employee Status', 'Records Status', 'Break Time']
    df = df[columns_order]

    def remove_first_in_last_out(records):
        entries = records.split(', ')
        if len(entries) > 0 and entries[0].endswith('in(ED)'):
            entries.pop(0)
        if len(entries) > 0 and entries[-1].endswith('out(ED)'):
            entries.pop(-1)
        return ', '.join(entries)

    # Apply the function to the 'Records' column
    df['Records'] = df['Records'].apply(lambda x: remove_first_in_last_out(x) if pd.notna(x) else x)

    def calculate_break_time(row):
        if row['Employee Status'] == 'Absent':
            return 'N/A'

        if row['Employee Status'] == 'Present':
            entries = row['Records'].split(',')
            total_break_time = 0
            for i in range(1, len(entries), 2):
                in_time_str = entries[i - 1].split()[-1]
                out_time_str = entries[i].split()[-1]

                in_time_match = re.search(r'\d{2}:\d{2}', in_time_str)
                out_time_match = re.search(r'\d{2}:\d{2}', out_time_str)

                if in_time_match and out_time_match:
                    in_time = pd.to_datetime(in_time_match.group(), format='%H:%M')
                    out_time = pd.to_datetime(out_time_match.group(), format='%H:%M')
                    break_duration = out_time - in_time
                    total_break_time += break_duration.total_seconds() / 60

                if row['Records Status'] == 'Punch records missing':
                    return 'N/A'

            return int(total_break_time)
        return 0

    def format_break_time(minutes):
        if minutes == 'N/A':
            return minutes
        hours = minutes // 60
        mins = minutes % 60
        if hours > 0:
            return f"{hours} hr {mins} mins" if mins > 0 else f"{hours} hr"
        else:
            return f"{mins} mins"

    # Apply the function to the DataFrame
    df['Break Time'] = df.apply(calculate_break_time, axis=1)

    # Format 'Break Time (minutes)' to the desired string format
    df['Break Time'] = df['Break Time'].apply(format_break_time)

    df['Punch Records'].replace('NaN', pd.NA, inplace=True)
    df['Records'].replace('NaN', pd.NA, inplace=True)

    # Adding the "Status" column based on the "Records" column
    df['Employee Status'] = df['Records'].apply(lambda x: 'Present' if pd.notna(x) and x != '' else 'Absent')

    # List of columns to drop
    columns_to_drop = ['Records']

    # Strip leading and trailing whitespaces from column names
    df.columns = df.columns.str.strip()

    # Drop the specified columns
    df.drop(columns=columns_to_drop, errors='ignore', inplace=True)

    # Function to mark columns empty based on specific words
    def mark_columns_empty(row):
        words_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                        'Punch Records', 'Emp Code:', 'Employee Name :', 'Department:', 'Records']
        
        if row.name > 0:  # Check if row index is greater than 0 (excluding headers)
            for word in words_to_check:
                if word in row and pd.notna(row[word]) and any(word in str(cell) for cell in row):
                    row['Employee Status'] = " "
                    row['Records Status'] = " "
                    row['Break Time'] = " "
                    return row
        return row

    # Function to update 'Records Status' column
    def update_status_based_on_records(records, punch_records):
        if pd.isna(records) or records.strip() == '':
            return 'Absent'

        entries = records.split(', ')

        if len(entries) == 1:
            return 'Punch records missing'

        if 'out' in entries[0]:
            return 'Punch records missing'

        if len(entries) % 2 != 0:
            return 'Punch records missing'

        # Check for consecutive 'in' or 'out' entries
        for i in range(1, len(entries)):
            if ('in' in entries[i] and 'in' in entries[i-1]) or \
            ('out' in entries[i] and 'out' in entries[i-1]):
                return 'Punch records missing'

        return 'Valid Records'

    # Apply the function to update 'Records Status' column
    if 'Records' in df.columns and 'Punch Records' in df.columns:
        df['Records Status'] = df.apply(lambda row: update_status_based_on_records(row['Records'], row['Punch Records']), axis=1)

    # Columns to check for NaN or words and alphabets
    cols_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                    'Punch Records', 'Emp Code:', 'Employee Name :', 'Department:', 'Records']

    def check_nan_and_update_status(row, cols_to_check):
        # If the entire row is NaN, set 'Employee Status', 'Records Status' and 'Break Time (minutes)' to empty space
        if row.isna().all():
            row['Employee Status'] = " "
            row['Records Status'] = " "
            row['Break Time'] = " "
        # If 'Records' column is empty, set 'Employee Status' and 'Break Time (minutes)' to empty space
        elif 'Records' in row and (pd.isna(row['Records']) or row['Records'].strip() == ''):
            row['Employee Status'] = " "
            row['Break Time'] = " "
        # If 'Records Status' is empty, set 'Employee Status' and 'Break Time (minutes)' to empty space
        elif 'Records Status' in row and row['Records Status'].strip() == '':
            row['Employee Status'] = " "
            row['Break Time'] = " "
        else:
            if 'Punch Records' in row and pd.notna(row['Punch Records']) and row['Punch Records'].strip() != '':
                row['Employee Status'] = "Present"
            else:
                row['Employee Status'] = "Absent"
        return row

    # Apply the function to update 'Employee Status', 'Records Status', and 'Break Time (minutes)' columns
    df = df.apply(lambda row: check_nan_and_update_status(row, cols_to_check), axis=1)

    # Apply the function to mark columns empty based on specific words
    df = df.apply(mark_columns_empty, axis=1)

    # Reorder columns to ensure "Employee Status", "Records Status", and "Break Time (minutes)" are last
    columns_order = [col for col in df.columns if col not in ['Employee Status', 'Records Status', 'Break Time']]
    columns_order += ['Employee Status', 'Records Status', 'Break Time']
    df = df[columns_order]

    def should_drop_row(row):
        first_cell_value = str(row.iloc[0])
        return first_cell_value.startswith(('Total', 'Department', 'Emp Code'))

    # Apply the function to filter out rows
    df = df[~df.apply(should_drop_row, axis=1)]

    # Create a temporary Excel file to store the formatted data
    temp_file_path = 'temp_file.xlsx'
    df.to_excel(temp_file_path, index=False)

    # Load the workbook and get the active worksheet
    wb = load_workbook(temp_file_path)
    ws = wb.active

    # Define fill colors
    aqua_fill = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")

    # Get the index of the "Break Time (minutes)" column
    break_time_col_idx = df.columns.get_loc('Break Time') + 1  # openpyxl uses 1-based indexing

    # Apply aqua color to the "Break Time (minutes)" column
    for row in ws.iter_rows(min_row=2, min_col=break_time_col_idx, max_col=break_time_col_idx):
        for cell in row:
            cell.fill = aqua_fill

    # Save the formatted workbook to a new file
    wb.save('formatted_file.xlsx')

    # Prepare the file for download
    output = "formatted_file.xlsx"

    # Send the file as an attachment
    return send_file(output, download_name="formatted_file.xlsx", as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True)