
import re
import pandas as pd
import json
import datetime
from openpyxl import load_workbook

# Global variables for system id, date, and SEQ per sheet
now = datetime.datetime.now()
systemid_value = f"{now.hour:02d}{now.minute:02d}{now.second:02d}"
system_date_value = now.strftime('%Y-%m-%d')
# seq_per_sheet_dict: {sheet_index: SEQ}
seq_per_sheet_dict = {}

# Mapping for MAPPING value (Excel cell value -> mapped number)
MAPPING_VALUE_DICT = {
    '項目定義書_帳票': '1',
    '項目定義書_画面': '2',
    '項目定義書_CSV': '3',
    '項目定義書_IPO図': '4',
    '項目定義書_ﾒﾆｭｰ': '5',
}

# Mapping for KOUMOKU types
KOUMOKU_TYPE_MAPPING = {
    'ラベル': '1',
    'チェックボックス': '2',
    '処理': '3',
}

# Dictionary to store SEQ_K and row numbers for each sheet
seq_k_per_sheet_dict = {}  # {sheet_index: {row_number: seq_k}}

# Global stop values for process_koumoku_data (excluding '【項目定義】')
STOP_VALUES = {
    '【帳票データ】',
    '【ファンクション定義】',
    '【メッセージ定義】',
    '【タブインデックス定義】',
    '【CSVデータ】',
    '【備考】',
    '【運用上の注意点】',
    '【項目定義】'
}

def read_table_info_to_dict(filename):
    """
    Reads the JSON content from the given filename and returns it as a dictionary.
    """
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"Keys in data: {list(data.keys())}")
    return data


def get_cell_value_with_merged(ws, cell_ref):
    """Helper function to get cell value considering merged cells"""
    cell = ws[cell_ref]
    if cell.value is not None:
        return cell.value
    # If cell is empty, check merged cells
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws[merged_range.start_cell.coordinate].value
    return None


def process_column_value_koumoku(col_info, ws, row_num, sheet_seq, seq_k_value, seq_k_l_value=None):
    """Process column value for T_KIHON_PJ_KOUMOKU table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_K':
        val = str(seq_k_value) if seq_k_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_k_value) if seq_k_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'SEQ_K_L':
        val = str(seq_k_l_value) if seq_k_l_value is not None else "''"
    elif val_rule == 'MAPPING':
        if cell_fix:
            # Get column letter from cell_fix (e.g., 'B' from 'B')
            col_letter = cell_fix
            cell_ref = f"{col_letter}{row_num}"
            cell_value = ws[cell_ref].value if ws[cell_ref].value else None
            val = KOUMOKU_TYPE_MAPPING.get(cell_value, "''")
        else:
            val = "''"
    elif val_rule == '':
        if cell_fix:
            try:
                col_letter = cell_fix
                cell_ref = f"{col_letter}{row_num}"
                cell_value = get_cell_value_with_merged(ws, cell_ref)
                if cell_value is None:
                    val = "''"
                elif isinstance(cell_value, str):
                    if col_info.get('DATA_TYPE', '').lower() == 'nvarchar':
                        val = f"N'{cell_value}'"
                    else:
                        val = f"'{cell_value}'"
                elif isinstance(cell_value, (int, float)):
                    val = str(cell_value)
                elif isinstance(cell_value, datetime.datetime):
                    val = f"'{cell_value.strftime('%Y-%m-%d %H:%M:%S')}'"
                else:
                    val = f"'{str(cell_value)}'"
            except Exception:
                val = "''"
        else:
            val = "''"
    elif val_rule == 'T_KIHON_PJ_GAMEN.SEQ':
        val = str(sheet_seq) if sheet_seq is not None else "''"
    elif val_rule == 'T_KIHON_PJ_KOUMOKU.SEQ_K':
        val = str(seq_k_value) if seq_k_value is not None else "''"
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value(col_info, ws, systemid_value, system_date_value, seq_value=None, jyun_value=None):
    """Process column value based on VALUE rules"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'BLANK':
        val = "''"
    elif val_rule == 'NULL':
        val = "NULL"
    elif val_rule == 'SYSTEMID':
        val = f"'{systemid_value}'"
    elif val_rule == 'T_KIHON_PJ.SYSTEM_ID':
        val = f"'{systemid_value}'"
    elif val_rule == 'AUTO_ID' and col_name == 'SEQ':
        val = str(seq_value) if seq_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'JYUN':
        val = str(jyun_value) if jyun_value is not None else "''"
    elif val_rule in ('SYSTEM DATE', 'AUTO_TIME'):
        val = f"'{system_date_value}'"
    elif val_rule == 'MAPPING':
        cell_value = ws[cell_fix].value if cell_fix else None
        val = MAPPING_VALUE_DICT.get(cell_value, "''")
    elif val_rule == '':
        if cell_fix:
            try:
                cell_value = get_cell_value_with_merged(ws, cell_fix)
                if cell_value is None:
                    val = "''"
                elif isinstance(cell_value, str):
                    # Add N prefix for nvarchar columns
                    if col_info.get('DATA_TYPE', '').lower() == 'nvarchar':
                        val = f"N'{cell_value}'"
                    else:
                        val = f"'{cell_value}'"
                elif isinstance(cell_value, (int, float)):
                    val = str(cell_value)
                elif isinstance(cell_value, datetime.datetime):
                    val = f"'{cell_value.strftime('%Y-%m-%d %H:%M:%S')}'"
                else:
                    val = f"'{str(cell_value)}'"
            except Exception:
                val = "''"
        else:
            val = "''"
    else:
        # Other values, treat as string literal
        # Add N prefix for nvarchar columns
        if col_info.get('DATA_TYPE', '').lower() == 'nvarchar':
            val = f"N'{val_rule}'"
        else:
            val = f"'{val_rule}'"
    
    return val
  


def generate_insert_statements_from_excel(excel_file, sheet_index, table_info_file, table_key):
    """
    Unified function to generate INSERT statements for all table types
    """
    # Read table info JSON
    table_info = read_table_info_to_dict(table_info_file)
    if table_key not in table_info:
        raise ValueError(f"Table key '{table_key}' not found in table info.")
    
    columns_info = table_info[table_key]
    insert_statements = []
    
    # Use global systemid_value and system_date_value
    global systemid_value, system_date_value
    
    if table_key == 'T_KIHON_PJ_GAMEN':
        # Special handling for T_KIHON_PJ_GAMEN: process multiple sheets
        global seq_per_sheet_dict
        wb = load_workbook(excel_file, data_only=True)
        sheetnames = wb.sheetnames
        seq_per_sheet = 1
        allowed_b2_values = set(MAPPING_VALUE_DICT.keys())
        for sheet_idx in range(2, len(sheetnames)):
            ws = wb[sheetnames[sheet_idx]]
            try:
                sheet_check_value = ws["B2"].value
            except Exception:
                sheet_check_value = None
            if sheet_check_value not in allowed_b2_values:
                continue
            row_data = {}
            seq_value = seq_per_sheet
            jyun_value = seq_value
            seq_per_sheet_dict[sheet_idx] = seq_value
            for col_info in columns_info:
                col_name = col_info.get('COLUMN_NAME', '')
                val = process_column_value(col_info, ws, systemid_value, system_date_value, seq_value, jyun_value)
                row_data[col_name] = val
            columns_str = ", ".join(row_data.keys())
            values_str = ", ".join(row_data.values())
            sql = f"INSERT INTO {table_key} ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
            seq_per_sheet += 1
    
    elif table_key == 'T_KIHON_PJ':
        # Special handling for T_KIHON_PJ: single insert statement
        wb = load_workbook(excel_file, data_only=True)
        sheetnames = wb.sheetnames
        if sheet_index >= len(sheetnames):
            raise ValueError(f"Sheet index {sheet_index} out of range.")
        ws = wb[sheetnames[sheet_index]]
        
        cols = []
        vals = []
        for col_info in columns_info:
            col_name = col_info['COLUMN_NAME']
            cols.append(col_name)
            val = process_column_value(col_info, ws, systemid_value, system_date_value)
            vals.append(val)

        columns_str = ", ".join(cols)
        values_str = ", ".join(vals)
        sql = f"INSERT INTO {table_key} ({columns_str}) VALUES ({values_str});"
        insert_statements.append(sql)
    
    else:
        # Default handling for other tables: process each row in the sheet
        df = pd.read_excel(excel_file, sheet_name=sheet_index, engine='openpyxl')
        wb = load_workbook(excel_file, data_only=True)
        sheetnames = wb.sheetnames
        if sheet_index >= len(sheetnames):
            raise ValueError(f"Sheet index {sheet_index} out of range.")
        ws = wb[sheetnames[sheet_index]]
        
        for _, row in df.iterrows():
            cols = []
            vals = []
            for col_info in columns_info:
                col_name = col_info['COLUMN_NAME']
                cols.append(col_name)
                val = process_column_value(col_info, ws, systemid_value, system_date_value)
                vals.append(val)
            columns_str = ", ".join(cols)
            values_str = ", ".join(vals)
            sql = f"INSERT INTO {table_key} ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
    
    return insert_statements


def process_all_tables_in_sequence(excel_file, table_info_file, output_file='insert_all.sql'):
    """
    Process all tables in the correct sequence:
    1. Create INSERT for T_KIHON_PJ
    2. Iterate through sheets (from sheet 3) to create INSERT for T_KIHON_PJ_GAMEN
    3. For each new SEQ, process T_KIHON_PJ_KOUMOKU
    4. For each new SEQ_K, process T_KIHON_PJ_KOUMOKU_LOGIC
    """
    global seq_per_sheet_dict, seq_k_per_sheet_dict
    
    all_insert_statements = []
    
    # Step 1: Create INSERT for T_KIHON_PJ (using sheet 3)
    print("Processing T_KIHON_PJ...")
    pj_inserts = generate_insert_statements_from_excel(excel_file, 2, table_info_file, 'T_KIHON_PJ')
    all_insert_statements.extend(pj_inserts)
    
    # Step 2: Process T_KIHON_PJ_GAMEN for sheets from index 2 onwards
    print("Processing T_KIHON_PJ_GAMEN...")
    table_info = read_table_info_to_dict(table_info_file)
    gamen_columns_info = table_info.get('T_KIHON_PJ_GAMEN', [])
    
    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames
    seq_per_sheet = 1
    allowed_b2_values = set(MAPPING_VALUE_DICT.keys())
    
    for sheet_idx in range(2, len(sheetnames)):
        ws = wb[sheetnames[sheet_idx]]
        try:
            sheet_check_value = ws["B2"].value
        except Exception:
            sheet_check_value = None
        
        if sheet_check_value not in allowed_b2_values:
            continue
        
        # Create T_KIHON_PJ_GAMEN insert
        row_data = {}
        seq_value = seq_per_sheet
        jyun_value = seq_value
        seq_per_sheet_dict[sheet_idx] = seq_value
        
        for col_info in gamen_columns_info:
            col_name = col_info.get('COLUMN_NAME', '')
            val = process_column_value(col_info, ws, systemid_value, system_date_value, seq_value, jyun_value)
            row_data[col_name] = val
        
        columns_str = ", ".join(row_data.keys())
        values_str = ", ".join(row_data.values())
        sql = f"INSERT INTO T_KIHON_PJ_GAMEN ({columns_str}) VALUES ({values_str});"
        all_insert_statements.append(sql)
        
        print(f"Processing sheet {sheet_idx}: {sheetnames[sheet_idx]} with SEQ {seq_value}")
        
        # Step 3: Process T_KIHON_PJ_KOUMOKU for this sheet
        koumoku_inserts = process_koumoku_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(koumoku_inserts)
        
        seq_per_sheet += 1
    
    # Write all statements to file
    with open(output_file, 'w', encoding='utf-8') as f:
        for stmt in all_insert_statements:
            f.write(stmt + '\n')
    
    print(f"All INSERT statements written to {output_file}")
    return all_insert_statements


def process_koumoku_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【項目定義】'
):
    """
    Process KOUMOKU data for a single sheet
    Returns list of INSERT statements for both T_KIHON_PJ_KOUMOKU and T_KIHON_PJ_KOUMOKU_LOGIC
    """
    global seq_k_per_sheet_dict
    
    if stop_values is None:
        stop_values = STOP_VALUES
    
    table_info = read_table_info_to_dict(table_info_file)
    koumoku_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU', [])
    koumoku_logic_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU_LOGIC', [])
    
    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames
    
    if sheet_idx >= len(sheetnames):
        return []
    
    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_k_counter = 1
    seq_k_per_sheet_dict[sheet_idx] = {}
    current_seq_k = None
    
    print(f"  Processing KOUMOKU data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
    # Scan from top to bottom for cell_b_value
    for row_num in range(1, ws.max_row + 1):
        cell_b = ws[f"B{row_num}"]
        if cell_b.value == cell_b_value:
            # Check subsequent rows
            for check_row in range(row_num + 1, ws.max_row + 1):
                cell_b_check = ws[f"B{check_row}"].value
                # Skip if value is cell_b_value, break if in stop_values
                if cell_b_check == cell_b_value:
                    continue
                if cell_b_check in stop_values:
                    break
                
                # Check if B and C are merged and have value != '画面' and != '番号'
                merged_bc = False
                for merged_range in ws.merged_cells.ranges:
                    if f"B{check_row}" in merged_range and f"C{check_row}" in merged_range:
                        merged_bc = True
                        break
                
                if merged_bc:
                    cell_b_val = ws[f"B{check_row}"].value
                    if cell_b_val and cell_b_val != '画面' and cell_b_val != '番号':
                        # Create KOUMOKU insert
                        current_seq_k = seq_k_counter
                        seq_k_per_sheet_dict[sheet_idx][check_row] = current_seq_k
                        
                        row_data = {}
                        for col_info in koumoku_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_koumoku(col_info, ws, check_row, sheet_seq, current_seq_k)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created KOUMOKU with SEQ_K {current_seq_k} at row {check_row}")
                        
                        # Step 4: Process T_KIHON_PJ_KOUMOKU_LOGIC for this SEQ_K
                        logic_inserts = process_koumoku_logic_for_seq_k(
                            ws, check_row, sheet_seq, current_seq_k, koumoku_logic_columns_info
                        )
                        insert_statements.extend(logic_inserts)
                        
                        seq_k_counter += 1
    
    return insert_statements


def process_koumoku_logic_for_seq_k(ws, start_row, sheet_seq, seq_k_value, koumoku_logic_columns_info):
    """
    Process T_KIHON_PJ_KOUMOKU_LOGIC for a specific SEQ_K
    """
    insert_statements = []
    seq_k_l_counter = 1
    
    # Check if B~BN are merged (indicating KOUMOKU_LOGIC data)
    for check_row in range(start_row, ws.max_row + 1):
        merged_b_to_bn = False
        for merged_range in ws.merged_cells.ranges:
            if f"B{check_row}" in merged_range:
                # Check if range extends to at least BN (column 66)
                start_col = merged_range.min_col
                end_col = merged_range.max_col
                if start_col == 2 and end_col >= 66:  # B=2, BN=66
                    merged_b_to_bn = True
                    break
        
        if merged_b_to_bn:
            # Create KOUMOKU_LOGIC insert
            row_data = {}
            for col_info in koumoku_logic_columns_info:
                col_name = col_info.get('COLUMN_NAME', '')
                val = process_column_value_koumoku(col_info, ws, check_row, sheet_seq, seq_k_value, seq_k_l_counter)
                row_data[col_name] = val
            
            columns_str = ", ".join(row_data.keys())
            values_str = ", ".join(row_data.values())
            sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU_LOGIC ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
            
            print(f"      Created KOUMOKU_LOGIC with SEQ_K_L {seq_k_l_counter} at row {check_row}")
            seq_k_l_counter += 1
        
        # Stop if we hit a stop value or another KOUMOKU section
        cell_b_check = ws[f"B{check_row}"].value
        if cell_b_check in STOP_VALUES or cell_b_check == '【項目定義】':
            break
    
    return insert_statements


# Example usage:
print("Starting processing all tables in sequence...")
all_inserts = process_all_tables_in_sequence('doc.xlsx', 'table_info.txt', 'insert_all.sql')
print(f"Generated {len(all_inserts)} INSERT statements in total.")


