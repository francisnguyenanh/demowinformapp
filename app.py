import re
import pandas as pd
import json


def read_table_info_to_dict(filename):
    """
    Reads the JSON content from the given filename and returns it as a dictionary.
    """
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"Keys in data: {list(data.keys())}")
    return data


def generate_insert_statements_from_excel(excel_file, sheet_index, table_info_file, table_key):
    """
    Reads the specified sheet from the Excel file and generates SQL Server INSERT statements
    based on the columns defined in table_info_file under table_key.
    The VALUE field is set according to specific rules.
    """
    import datetime

    # Read the sheet from Excel with openpyxl engine to access merged cells
    df = pd.read_excel(excel_file, sheet_name=sheet_index, engine='openpyxl')

    # Load workbook and sheet to handle merged cells
    from openpyxl import load_workbook
    wb = load_workbook(excel_file, data_only=True)
    sheetnames = wb.sheetnames
    if sheet_index >= len(sheetnames):
        raise ValueError(f"Sheet index {sheet_index} out of range.")
    ws = wb[sheetnames[sheet_index]]

    # Read table info JSON
    table_info = read_table_info_to_dict(table_info_file)

    if table_key not in table_info:
        raise ValueError(f"Table key '{table_key}' not found in table info.")

    columns_info = table_info[table_key]
    column_names = [col['COLUMN_NAME'] for col in columns_info]

    # Helper to get cell value considering merged cells
    def get_cell_value(cell_ref):
        cell = ws[cell_ref]
        if cell.value is not None:
            return cell.value
        # If cell is empty, check merged cells
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws[merged_range.start_cell.coordinate].value
        return None

    # Prepare insert statements list
    insert_statements = []

    # Current system time and date for SYSTEMID and SYSTEM DATE/AUTO_TIME
    now = datetime.datetime.now()
    systemid_value = f"{now.hour:02d}{now.minute:02d}{now.second:02d}"
    system_date_value = now.strftime('%Y-%m-%d')

    # For each row in the dataframe, generate an INSERT statement
    # But for key 'T_KIHON_PJ', only output one insert statement
    if table_key == 'T_KIHON_PJ':
        cols = []
        vals = []
        for col_info in columns_info:
            col_name = col_info['COLUMN_NAME']
            cols.append(col_name)
            val_rule = col_info.get('VALUE', '')
            cell_fix = col_info.get('CELL_FIX', '').strip()

            if val_rule == 'BLANK':
                val = "''"
            elif val_rule == 'NULL':
                val = "NULL"
            elif val_rule == 'SYSTEMID':
                val = f"'{systemid_value}'"
            elif val_rule in ('SYSTEM DATE', 'AUTO_TIME'):
                val = f"'{system_date_value}'"
            elif val_rule == '':
                if cell_fix:
                    try:
                        cell_value = get_cell_value(cell_fix)
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

            vals.append(val)

        columns_str = ", ".join(cols)
        values_str = ", ".join(vals)
        sql = f"INSERT INTO {table_key} ({columns_str}) VALUES ({values_str});"
        insert_statements.append(sql)
    else:
        for _, row in df.iterrows():
            cols = []
            vals = []
            for col_info in columns_info:
                col_name = col_info['COLUMN_NAME']
                cols.append(col_name)
                val_rule = col_info.get('VALUE', '')
                cell_fix = col_info.get('CELL_FIX', '').strip()

                if val_rule == 'BLANK':
                    val = "''"
                elif val_rule == 'NULL':
                    val = "NULL"
                elif val_rule == 'SYSTEMID':
                    val = f"'{systemid_value}'"
                elif val_rule in ('SYSTEM DATE', 'AUTO_TIME'):
                    val = f"'{system_date_value}'"
                elif val_rule == '':
                    if cell_fix:
                        try:
                            cell_value = get_cell_value(cell_fix)
                            if cell_value is None:
                                val = "''"
                            elif isinstance(cell_value, str):
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
                    val = f"'{val_rule}'"

                vals.append(val)

            columns_str = ", ".join(cols)
            values_str = ", ".join(vals)
            sql = f"INSERT INTO {table_key} ({columns_str}) VALUES ({values_str});"
            insert_statements.append(sql)

    return insert_statements


# Example usage:
insert_stmts = generate_insert_statements_from_excel('doc.xlsx', 2, 'table_info.txt', 'T_KIHON_PJ')
with open('insert.sql', 'w', encoding='utf-8') as f:
    for stmt in insert_stmts:
        f.write(stmt + '\n')


