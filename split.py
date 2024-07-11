import pandas as pd

def split_tables(df):
    tables = []
    current_table = None
    current_headers = []

    for idx, row in df.iterrows():
        if pd.isna(row[0]) and pd.isna(row[1]):
            if current_table is not None:
                tables.append((current_headers, current_table))
            current_headers = [row.tolist()]
            current_table = pd.DataFrame()
        elif current_headers and not pd.isna(row[0]):
            current_table = pd.concat([current_table, pd.DataFrame([row.tolist()])], ignore_index=True)
        else:
            current_headers.append(row.tolist())

    if current_table is not None:
        tables.append((current_headers, current_table))
    
    return tables
