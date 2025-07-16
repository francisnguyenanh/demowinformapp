import re
import pandas as pd
import json

def try_readlines(filename, encodings):
    for enc in encodings:
        try:
            with open(filename, encoding=enc) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"Cannot decode {filename} with tried encodings: {encodings}")

def read_excel_to_json(excel_file, output_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Replace 'blank' string in 'VALUE' column with empty string ""
    if 'VALUE' in df.columns:
        df['VALUE'] = df['VALUE'].replace('BLANK', '')

    # Replace NaN with empty string in entire dataframe
    df = df.fillna('')

    # Convert CHARACTER_MAXIMUM_LENGTH to integer if possible
    if 'CHARACTER_MAXIMUM_LENGTH' in df.columns:
        def to_int_or_none(x):
            try:
                return int(x)
            except (ValueError, TypeError):
                return x
        df['CHARACTER_MAXIMUM_LENGTH'] = df['CHARACTER_MAXIMUM_LENGTH'].apply(to_int_or_none)

    # Group by TABLE_NAME
    grouped = {}
    for table_name, group_df in df.groupby('TABLE_NAME'):
        grouped[table_name] = group_df.drop(columns=['TABLE_NAME']).to_dict(orient='records')

    # Define the desired key order
    key_order = ['T_KIHON_PJ', 'T_KIHON_PJ_GAMEN', 'T_KIHON_PJ_GAMEN_YOUKEN', 'T_KIHON_PJ_KOUMOKU', 'T_KIHON_PJ_KOUMOKU_LOGIC', 'T_KIHON_PJ_FUNC', 'T_KIHON_PJ_FUNC_LOGIC', 'T_KIHON_PJ_IPO',  'T_KIHON_PJ_KOUMOKU_CSV', 'T_KIHON_PJ_KOUMOKU_CSV_LOGIC',  'T_KIHON_PJ_KOUMOKU_RE', 'T_KIHON_PJ_KOUMOKU_RE_LOGIC', 'T_KIHON_PJ_MENU', 'T_KIHON_PJ_MESSAGE', 'T_KIHON_PJ_TAB']

    # Create an ordered dictionary based on the key order
    from collections import OrderedDict
    ordered_grouped = OrderedDict()
    for key in key_order:
        if key in grouped:
            ordered_grouped[key] = grouped[key]

    # Convert to JSON string with indentation for readability
    json_str = json.dumps(ordered_grouped, ensure_ascii=False, indent=4)

    # Write JSON string to output file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(json_str)

# Existing code to read input.txt and extract unique table names
encodings = ['utf-8', 'utf-8-sig', 'cp1252', 'latin-1']
lines = try_readlines('input.txt', encodings)

table_names = []

# Chỉ focus vào các câu lệnh INSERT
pattern = re.compile(r'^\s*INSERT INTO\s+([A-Z0-9_]+)\s*\(', re.IGNORECASE)

for line in lines:
    match = pattern.search(line)
    if match:
        table_names.append(match.group(1))

# Loại bỏ trùng lặp, giữ nguyên thứ tự xuất hiện
unique_table_names = list(dict.fromkeys(table_names))

# Ghi danh sách bảng ra file output.txt, mỗi bảng trên một dòng
with open('TABLE_INFO.txt', 'w', encoding='utf-8') as f:
    for name in unique_table_names:
        f.write(name + '\n')

print(unique_table_names)

# New code to read mapping.xlsx and write JSON to output.txt
read_excel_to_json('mapping.xlsx', 'TABLE_INFO.txt')
