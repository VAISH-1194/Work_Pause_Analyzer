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

        df_non_empty = df_initial.dropna(how='all')

        first_non_empty_row = df_non_empty.index[0]

        df = pd.read_excel(file, header=first_non_empty_row)

        df = pd.read_excel(file, skiprows=9)

        df = df.drop([col for col in df.columns if col.startswith('Unnamed')], axis=1)

        df['Records'] = df['Punch Records']

    def change_name(records):
        if pd.isna(records):
            return records

        entries = records.split(',')
        entries = [entry.replace('BD', 'ED') for entry in entries]

        entries = [entry.replace('Main Entrance', 'ED').replace('Exit', 'ED') for entry in entries]

        return ', '.join(entries)

    df['Records'] = df['Records'].apply(change_name)

    def filter_punch_records(record):
        if pd.isna(record):
            return record

        entries = record.split(',')
        valid_entries = [entry for entry in entries if ('in' in entry or 'out' in entry)]

        return ','.join(valid_entries)

    df['Records'] = df['Records'].apply(filter_punch_records)

    df['Punch Records'].replace('NaN', pd.NA, inplace=True)
    df['Records'].replace('NaN', pd.NA, inplace=True)

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

        for i in range(1, len(entries)):
            if ('in' in entries[i] and 'in' in entries[i-1]) or \
            ('out' in entries[i] and 'out' in entries[i-1]):
                return 'Punch records missing'

        return 'Valid Records'

    def mark_columns_empty(row):
        words_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                        'Punch Records', 'Records']
        
        if row.name > 0: 
            for word in words_to_check:
                if pd.notna(row[word]) and any(word in str(cell) for cell in row):
                    row['Employee Status'] = " "
                    row['Records Status'] = " "
                    row['Break Time'] = " "
                    return row
        return row

    df = df.apply(mark_columns_empty, axis=1)

    df['Records Status'] = df.apply(lambda row: update_status_based_on_records(row['Records'], row['Punch Records']), axis=1)

    cols_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                    'Punch Records', 'Records']

    def check_nan_and_update_status(row, cols_to_check):
        if all(pd.isna(row[col]) for col in cols_to_check):
            row['Employee Status'] = " "
            row['Records Status'] = " "
            row['Break Time'] = " "
        return row

    df = df.apply(lambda row: check_nan_and_update_status(row, cols_to_check), axis=1)

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

    df['Break Time'] = df.apply(calculate_break_time, axis=1)

    df['Break Time'] = df['Break Time'].apply(format_break_time)

    df['Punch Records'].replace('NaN', pd.NA, inplace=True)
    df['Records'].replace('NaN', pd.NA, inplace=True)

    df['Employee Status'] = df['Records'].apply(lambda x: 'Present' if pd.notna(x) and x != '' else 'Absent')

    columns_to_drop = ['Records']

    df.columns = df.columns.str.strip()

    df.drop(columns=columns_to_drop, errors='ignore', inplace=True)

    def mark_columns_empty(row):
        words_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                        'Punch Records', 'Records']
        
        if row.name > 0:
            for word in words_to_check:
                if word in row and pd.notna(row[word]) and any(word in str(cell) for cell in row):
                    row['Employee Status'] = " "
                    row['Records Status'] = " "
                    row['Break Time'] = " "
                    return row
        return row

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

        for i in range(1, len(entries)):
            if ('in' in entries[i] and 'in' in entries[i-1]) or \
            ('out' in entries[i] and 'out' in entries[i-1]):
                return 'Punch records missing'

        return 'Valid Records'

    if 'Records' in df.columns and 'Punch Records' in df.columns:
        df['Records Status'] = df.apply(lambda row: update_status_based_on_records(row['Records'], row['Punch Records']), axis=1)

    cols_to_check = ['Att. Date', 'InTime', 'OutTime', 'Shift', 'S. InTime', 'S. OutTime',
                    'Punch Records', 'Records']

    def check_nan_and_update_status(row, cols_to_check):

        if row.isna().all():
            row['Employee Status'] = " "
            row['Records Status'] = " "
            row['Break Time'] = " "

        elif 'Records' in row and (pd.isna(row['Records']) or row['Records'].strip() == ''):
            row['Employee Status'] = " "
            row['Break Time'] = " "

        elif 'Records Status' in row and row['Records Status'].strip() == '':
            row['Employee Status'] = " "
            row['Break Time'] = " "
        else:
            if 'Punch Records' in row and pd.notna(row['Punch Records']) and row['Punch Records'].strip() != '':
                row['Employee Status'] = "Present"
            else:
                row['Employee Status'] = "Absent"
        return row


    df = df.apply(lambda row: check_nan_and_update_status(row, cols_to_check), axis=1)


    df = df.apply(mark_columns_empty, axis=1)


    columns_order = [col for col in df.columns if col not in ['Employee Status', 'Records Status', 'Break Time']]
    columns_order += ['Employee Status', 'Records Status', 'Break Time']
    df = df[columns_order]

    def should_drop_row(row):
        first_cell_value = str(row.iloc[0])
        return first_cell_value.startswith(('Total', 'Department', 'Emp Code'))

# Drop rows based on the condition
    df = df[~df.apply(should_drop_row, axis=1)]

# Reset the index if needed
    df.reset_index(drop=True, inplace=True)



    df = df[~df.apply(should_drop_row, axis=1)]

    temp_file_path = 'temp_file.xlsx'
    df.to_excel(temp_file_path, index=False)

    wb = load_workbook(temp_file_path)
    ws = wb.active

    aqua_fill = PatternFill(start_color="C9DAF8", end_color="C9DAF8", fill_type="solid")

    break_time_col_idx = df.columns.get_loc('Break Time') + 1  

    for row in ws.iter_rows(min_row=2, min_col=break_time_col_idx, max_col=break_time_col_idx):
        for cell in row:
            cell.fill = aqua_fill

    wb.save('formatted_file.xlsx')

    output = "formatted_file.xlsx"

    return send_file(output, download_name="formatted_file.xlsx", as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True)