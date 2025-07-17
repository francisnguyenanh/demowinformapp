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



# Global stop values for process_koumoku_data (excluding '【項目定義】')
STOP_VALUES = {
    '【帳票データ】',
    '【ファンクション定義】',
    '【メッセージ定義】',
    '【タブインデックス定義】',
    '【CSVデータ】',
    '【備考】',
    '【運用上の注意点】',
    '【項目定義】', 
    '【一覧定義】',
    '【表示位置定義】'
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


def process_column_value_func(col_info, ws, row_num, sheet_seq, seq_f_value, seq_f_l_value=None):
    """Process column value for T_KIHON_PJ_FUNC table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_F':
        val = str(seq_f_value) if seq_f_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_f_value) if seq_f_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'SEQ_F_L':
        val = str(seq_f_l_value) if seq_f_l_value is not None else "''"
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
    elif val_rule == 'T_KIHON_PJ_FUNC.SEQ_F':
        val = str(seq_f_value) if seq_f_value is not None else "''"
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_csv(col_info, ws, row_num, sheet_seq, seq_csv_value, seq_csv_l_value=None):
    """Process column value for T_KIHON_PJ_KOUMOKU_CSV table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_CSV':
        val = str(seq_csv_value) if seq_csv_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_csv_value) if seq_csv_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'SEQ_CSV_L':
        val = str(seq_csv_l_value) if seq_csv_l_value is not None else "''"
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
    elif val_rule == 'T_KIHON_PJ_KOUMOKU_CSV.SEQ_CSV':
        val = str(seq_csv_value) if seq_csv_value is not None else "''"
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_re(col_info, ws, row_num, sheet_seq, seq_re_value, seq_re_l_value=None):
    """Process column value for T_KIHON_PJ_KOUMOKU_RE table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_RE':
        val = str(seq_re_value) if seq_re_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_re_value) if seq_re_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'SEQ_RE_L':
        val = str(seq_re_l_value) if seq_re_l_value is not None else "''"
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
    elif val_rule == 'T_KIHON_PJ_KOUMOKU_RE.SEQ_RE':
        val = str(seq_re_value) if seq_re_value is not None else "''"
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_message(col_info, ws, row_num, sheet_seq, seq_ms_value):
    """Process column value for T_KIHON_PJ_MESSAGE table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_MS':
        val = str(seq_ms_value) if seq_ms_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_ms_value) if seq_ms_value is not None else "''"
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
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_tab(col_info, ws, row_num, sheet_seq, seq_t_value):
    """Process column value for T_KIHON_PJ_TAB table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_T':
        val = str(seq_t_value) if seq_t_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_t_value) if seq_t_value is not None else "''"
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
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_ichiran(col_info, ws, row_num, sheet_seq, seq_i_value):
    """Process column value for T_KIHON_PJ_ICHIRAN table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_I':
        val = str(seq_i_value) if seq_i_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_i_value) if seq_i_value is not None else "''"
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
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_menu(col_info, ws, row_num, sheet_seq, seq_m_value):
    """Process column value for T_KIHON_PJ_MENU table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_M':
        val = str(seq_m_value) if seq_m_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_m_value) if seq_m_value is not None else "''"
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
    else:
        # Use existing process_column_value logic
        val = process_column_value(col_info, ws, systemid_value, system_date_value)
    
    return val


def process_column_value_ipo(col_info, ws, row_num, sheet_seq, seq_ipo_value):
    """Process column value for T_KIHON_PJ_IPO table"""
    val_rule = col_info.get('VALUE', '')
    cell_fix = col_info.get('CELL_FIX', '').strip()
    col_name = col_info.get('COLUMN_NAME', '')
    
    if val_rule == 'AUTO_ID' and col_name == 'SEQ_IPO':
        val = str(seq_ipo_value) if seq_ipo_value is not None else "''"
    elif val_rule == 'AUTO_ID' and col_name == 'ROW_NO':
        val = str(seq_ipo_value) if seq_ipo_value is not None else "''"
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
    global seq_per_sheet_dict
    
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
        
        # Step 4: Process T_KIHON_PJ_FUNC for this sheet
        func_inserts = process_func_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(func_inserts)
        
        # Step 5: Process T_KIHON_PJ_KOUMOKU_CSV for this sheet
        csv_inserts = process_csv_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(csv_inserts)
        
        # Step 6: Process T_KIHON_PJ_KOUMOKU_RE for this sheet
        re_inserts = process_re_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(re_inserts)
        
        # Step 7: Process T_KIHON_PJ_MESSAGE for this sheet
        message_inserts = process_message_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(message_inserts)
        
        # Step 8: Process T_KIHON_PJ_TAB for this sheet
        tab_inserts = process_tab_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(tab_inserts)

        # Step 9: Process T_KIHON_PJ_HYOUJI for this sheet
        hyouji_inserts = process_hyouji_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(hyouji_inserts)

        # Step 10: Process T_KIHON_PJ_ICHIRAN for this sheet
        ichiran_inserts = process_ichiran_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(ichiran_inserts)
        
        # Step 11: Process T_KIHON_PJ_MENU for this sheet
        menu_inserts = process_menu_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(menu_inserts)
        
        # Step 12: Process T_KIHON_PJ_IPO for this sheet
        ipo_inserts = process_ipo_data_for_single_sheet(
            excel_file, sheet_idx, seq_value, table_info_file
        )
        all_insert_statements.extend(ipo_inserts)
        
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


def process_func_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【ファンクション定義】'
):
    """
    Process FUNC data for a single sheet
    Returns list of INSERT statements for both T_KIHON_PJ_FUNC and T_KIHON_PJ_FUNC_LOGIC
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    func_columns_info = table_info.get('T_KIHON_PJ_FUNC', [])
    func_logic_columns_info = table_info.get('T_KIHON_PJ_FUNC_LOGIC', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_f_counter = 1
    current_seq_f = None

    print(f"  Processing FUNC data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create FUNC insert
                        current_seq_f = seq_f_counter
                        
                        row_data = {}
                        for col_info in func_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_func(col_info, ws, check_row, sheet_seq, current_seq_f)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_FUNC ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created FUNC with SEQ_F {current_seq_f} at row {check_row}")
                        
                        # Step 5: Process T_KIHON_PJ_FUNC_LOGIC for this SEQ_F
                        logic_inserts = process_func_logic_for_seq_f(
                            ws, check_row, sheet_seq, current_seq_f, func_logic_columns_info
                        )
                        insert_statements.extend(logic_inserts)
                        
                        seq_f_counter += 1
    
    return insert_statements


def process_func_logic_for_seq_f(ws, start_row, sheet_seq, seq_f_value, func_logic_columns_info):
    """
    Process T_KIHON_PJ_FUNC_LOGIC for a specific SEQ_F
    """
    insert_statements = []
    seq_f_l_counter = 1
    
    # Check if B~BN are merged (indicating FUNC_LOGIC data)
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
            # Create FUNC_LOGIC insert
            row_data = {}
            for col_info in func_logic_columns_info:
                col_name = col_info.get('COLUMN_NAME', '')
                val = process_column_value_func(col_info, ws, check_row, sheet_seq, seq_f_value, seq_f_l_counter)
                row_data[col_name] = val
            
            columns_str = ", ".join(row_data.keys())
            values_str = ", ".join(row_data.values())
            sql = f"INSERT INTO T_KIHON_PJ_FUNC_LOGIC ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
            
            print(f"      Created FUNC_LOGIC with SEQ_F_L {seq_f_l_counter} at row {check_row}")
            seq_f_l_counter += 1
        
        # Stop if we hit a stop value or another FUNC section
        cell_b_check = ws[f"B{check_row}"].value
        if cell_b_check in STOP_VALUES or cell_b_check == '【ファンクション定義】':
            break
    
    return insert_statements


def process_csv_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【CSVデータ】'
):
    """
    Process CSV data for a single sheet
    Returns list of INSERT statements for both T_KIHON_PJ_KOUMOKU_CSV and T_KIHON_PJ_KOUMOKU_CSV_LOGIC
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    csv_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU_CSV', [])
    csv_logic_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU_CSV_LOGIC', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_csv_counter = 1
    current_seq_csv = None

    print(f"  Processing CSV data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create CSV insert
                        current_seq_csv = seq_csv_counter
                        
                        row_data = {}
                        for col_info in csv_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_csv(col_info, ws, check_row, sheet_seq, current_seq_csv)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU_CSV ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created CSV with SEQ_CSV {current_seq_csv} at row {check_row}")
                        
                        # Step 6: Process T_KIHON_PJ_KOUMOKU_CSV_LOGIC for this SEQ_CSV
                        logic_inserts = process_csv_logic_for_seq_csv(
                            ws, check_row, sheet_seq, current_seq_csv, csv_logic_columns_info
                        )
                        insert_statements.extend(logic_inserts)
                        
                        seq_csv_counter += 1
    
    return insert_statements


def process_csv_logic_for_seq_csv(ws, start_row, sheet_seq, seq_csv_value, csv_logic_columns_info):
    """
    Process T_KIHON_PJ_KOUMOKU_CSV_LOGIC for a specific SEQ_CSV
    """
    insert_statements = []
    seq_csv_l_counter = 1
    
    # Check if B~BN are merged (indicating CSV_LOGIC data)
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
            # Create CSV_LOGIC insert
            row_data = {}
            for col_info in csv_logic_columns_info:
                col_name = col_info.get('COLUMN_NAME', '')
                val = process_column_value_csv(col_info, ws, check_row, sheet_seq, seq_csv_value, seq_csv_l_counter)
                row_data[col_name] = val
            
            columns_str = ", ".join(row_data.keys())
            values_str = ", ".join(row_data.values())
            sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU_CSV_LOGIC ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
            
            print(f"      Created CSV_LOGIC with SEQ_CSV_L {seq_csv_l_counter} at row {check_row}")
            seq_csv_l_counter += 1
        
        # Stop if we hit a stop value or another CSV section
        cell_b_check = ws[f"B{check_row}"].value
        if cell_b_check in STOP_VALUES or cell_b_check == '【CSVデータ】':
            break
    
    return insert_statements


def process_re_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【帳票データ】'
):
    """
    Process RE data for a single sheet
    Returns list of INSERT statements for both T_KIHON_PJ_KOUMOKU_RE and T_KIHON_PJ_KOUMOKU_RE_LOGIC
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    re_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU_RE', [])
    re_logic_columns_info = table_info.get('T_KIHON_PJ_KOUMOKU_RE_LOGIC', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_re_counter = 1
    current_seq_re = None

    print(f"  Processing RE data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create RE insert
                        current_seq_re = seq_re_counter
                        
                        row_data = {}
                        for col_info in re_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_re(col_info, ws, check_row, sheet_seq, current_seq_re)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU_RE ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created RE with SEQ_RE {current_seq_re} at row {check_row}")
                        
                        # Process T_KIHON_PJ_KOUMOKU_RE_LOGIC for this SEQ_RE
                        logic_inserts = process_re_logic_for_seq_re(
                            ws, check_row, sheet_seq, current_seq_re, re_logic_columns_info
                        )
                        insert_statements.extend(logic_inserts)
                        
                        seq_re_counter += 1
    
    return insert_statements


def process_re_logic_for_seq_re(ws, start_row, sheet_seq, seq_re_value, re_logic_columns_info):
    """
    Process T_KIHON_PJ_KOUMOKU_RE_LOGIC for a specific SEQ_RE
    """
    insert_statements = []
    seq_re_l_counter = 1
    
    # Check if B~BN are merged (indicating RE_LOGIC data)
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
            # Create RE_LOGIC insert
            row_data = {}
            for col_info in re_logic_columns_info:
                col_name = col_info.get('COLUMN_NAME', '')
                val = process_column_value_re(col_info, ws, check_row, sheet_seq, seq_re_value, seq_re_l_counter)
                row_data[col_name] = val
            
            columns_str = ", ".join(row_data.keys())
            values_str = ", ".join(row_data.values())
            sql = f"INSERT INTO T_KIHON_PJ_KOUMOKU_RE_LOGIC ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)
            
            print(f"      Created RE_LOGIC with SEQ_RE_L {seq_re_l_counter} at row {check_row}")
            seq_re_l_counter += 1
        
        # Stop if we hit a stop value or another RE section
        cell_b_check = ws[f"B{check_row}"].value
        if cell_b_check in STOP_VALUES or cell_b_check == '【帳票データ】':
            break
    
    return insert_statements


def process_message_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【メッセージ定義】'
):
    """
    Process MESSAGE data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_MESSAGE
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    message_columns_info = table_info.get('T_KIHON_PJ_MESSAGE', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_ms_counter = 1

    print(f"  Processing MESSAGE data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create MESSAGE insert
                        current_seq_ms = seq_ms_counter
                        
                        row_data = {}
                        for col_info in message_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_message(col_info, ws, check_row, sheet_seq, current_seq_ms)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_MESSAGE ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created MESSAGE with SEQ_MS {current_seq_ms} at row {check_row}")
                        
                        seq_ms_counter += 1
    
    return insert_statements


def process_tab_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【タブインデックス定義】'
):
    """
    Process TAB data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_TAB
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    tab_columns_info = table_info.get('T_KIHON_PJ_TAB', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_t_counter = 1

    print(f"  Processing TAB data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create TAB insert
                        current_seq_t = seq_t_counter
                        
                        row_data = {}
                        for col_info in tab_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_tab(col_info, ws, check_row, sheet_seq, current_seq_t)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_TAB ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created TAB with SEQ_T {current_seq_t} at row {check_row}")
                        
                        seq_t_counter += 1
    
    return insert_statements


def process_hyouji_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【表示位置定義】'
):
    """
    Process HYOUJI data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_HYOUJI
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    hyouji_columns_info = table_info.get('T_KIHON_PJ_HYOUJI', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_hyouji_counter = 1

    print(f"  Processing HYOUJI data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    # Scan from top to bottom for cell_b_value
    for row_num in range(1, ws.max_row + 1):
        cell_b = ws[f"B{row_num}"]
        if cell_b.value == cell_b_value:
            # Check subsequent rows
            for check_row in range(row_num + 1, ws.max_row + 1):
                cell_b_check = ws[f"B{check_row}"].value
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
                        current_seq_hyouji = seq_hyouji_counter
                        row_data = {}
                        for col_info in hyouji_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_message(col_info, ws, check_row, sheet_seq, current_seq_hyouji)
                            row_data[col_name] = val
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_HYOUJI (" + columns_str + ") VALUES (" + values_str + ");"
                        insert_statements.append(sql)
                        print(f"    Created HYOUJI with SEQ_HYOUJI {current_seq_hyouji} at row {check_row}")
                        seq_hyouji_counter += 1
    return insert_statements


def process_ichiran_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【一覧定義】'
):
    """
    Process ICHIRAN data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_ICHIRAN
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    ichiran_columns_info = table_info.get('T_KIHON_PJ_ICHIRAN', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_i_counter = 1

    print(f"  Processing ICHIRAN data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create ICHIRAN insert
                        current_seq_i = seq_i_counter
                        
                        row_data = {}
                        for col_info in ichiran_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_ichiran(col_info, ws, check_row, sheet_seq, current_seq_i)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_ICHIRAN ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created ICHIRAN with SEQ_I {current_seq_i} at row {check_row}")
                        
                        seq_i_counter += 1
    
    return insert_statements


def process_menu_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【メニュー定義】'
):
    """
    Process MENU data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_MENU
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    menu_columns_info = table_info.get('T_KIHON_PJ_MENU', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_m_counter = 1

    print(f"  Processing MENU data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create MENU insert
                        current_seq_m = seq_m_counter
                        
                        row_data = {}
                        for col_info in menu_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_menu(col_info, ws, check_row, sheet_seq, current_seq_m)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_MENU ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created MENU with SEQ_M {current_seq_m} at row {check_row}")
                        
                        seq_m_counter += 1
    
    return insert_statements


def process_ipo_data_for_single_sheet(
    excel_file, 
    sheet_idx, 
    sheet_seq, 
    table_info_file,
    stop_values=None,
    cell_b_value='【IPO定義】'
):
    """
    Process IPO data for a single sheet
    Returns list of INSERT statements for T_KIHON_PJ_IPO
    """
    if stop_values is None:
        stop_values = STOP_VALUES

    table_info = read_table_info_to_dict(table_info_file)
    ipo_columns_info = table_info.get('T_KIHON_PJ_IPO', [])

    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames

    if sheet_idx >= len(sheetnames):
        return []

    ws = wb[sheetnames[sheet_idx]]
    insert_statements = []
    seq_ipo_counter = 1

    print(f"  Processing IPO data for sheet {sheet_idx}: {sheetnames[sheet_idx]}")
    
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
                        # Create IPO insert
                        current_seq_ipo = seq_ipo_counter
                        
                        row_data = {}
                        for col_info in ipo_columns_info:
                            col_name = col_info.get('COLUMN_NAME', '')
                            val = process_column_value_ipo(col_info, ws, check_row, sheet_seq, current_seq_ipo)
                            row_data[col_name] = val
                        
                        columns_str = ", ".join(row_data.keys())
                        values_str = ", ".join(row_data.values())
                        sql = f"INSERT INTO T_KIHON_PJ_IPO ({columns_str}) VALUES ({values_str});"
                        insert_statements.append(sql)
                        
                        print(f"    Created IPO with SEQ_IPO {current_seq_ipo} at row {check_row}")
                        
                        seq_ipo_counter += 1
    
    return insert_statements


# Example usage:
print("Starting processing all tables in sequence...")
all_inserts = process_all_tables_in_sequence('doc.xlsx', 'table_info.txt', 'insert_all.sql')
print(f"Generated {len(all_inserts)} INSERT statements in total.")


