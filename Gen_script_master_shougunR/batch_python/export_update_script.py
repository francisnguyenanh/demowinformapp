import pyodbc
import os

# Đọc chuỗi kết nối từ file connect_string.txt
CONNECT_STRING_FILE = os.path.join(os.path.dirname(__file__), 'connect_string.txt')
def read_connect_string(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read().strip()

output_dir = 'output_scripts_update'  # Thư mục để lưu các file script UPDATE
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

conn_str = read_connect_string(CONNECT_STRING_FILE)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()


    # Lấy danh sách các bảng thực sự (BASE TABLE) bắt đầu bằng 'M_'
    cursor.execute("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME LIKE 'M_%'")
    tables = cursor.fetchall()

    for table in tables:
        table_name = table[0]
        # Lấy danh sách các cột có chứa các từ khóa cần thiết trong tên cột
        cursor.execute("""
            SELECT COLUMN_NAME 
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_NAME = ? AND (
                (COLUMN_NAME LIKE '%NAME%' OR
                COLUMN_NAME LIKE '%TEL%' OR
                COLUMN_NAME LIKE '%FAX%' OR
                COLUMN_NAME LIKE '%POST%' OR
                COLUMN_NAME LIKE '%ADDRESS%' OR
                COLUMN_NAME LIKE '%TANTOU%' OR
                COLUMN_NAME LIKE '%CREATE_USER%' OR
                COLUMN_NAME LIKE '%UPDATE_USER%' OR
                COLUMN_NAME LIKE '%FURIGANA%')
                AND COLUMN_NAME NOT LIKE '%CD%'
            )
        """, table_name)
        target_columns = [row[0] for row in cursor.fetchall()]

        if not target_columns:
            print(f"Bảng {table_name} không có cột nào chứa các từ khóa cần thiết")
            continue

        # Lấy dữ liệu các cột đó từ DB A, bỏ qua bảng bị lỗi
        select_query = f"SELECT {', '.join(target_columns)} FROM {table_name}"
        try:
            cursor.execute(select_query)
            rows = cursor.fetchall()
        except Exception as e:
            print(f"Bảng {table_name} bị lỗi khi truy vấn dữ liệu: {e}")
            continue

        # Sinh script UPDATE cho DB B
        script_file = os.path.join(output_dir, f"{table_name}_update.sql")
        with open(script_file, 'w', encoding='utf-8-sig') as f:  # UTF-8 with BOM
            f.write(f"-- UPDATE script for table {table_name} (columns containing NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA)\n")
            f.write(f"-- Số dòng dữ liệu: {len(rows)}\n\n")
            for idx, row in enumerate(rows):
                set_clauses = []
                for col, val in zip(target_columns, row):
                    if val is None:
                        set_clauses.append(f"{col} = NULL")
                    else:
                        # Escape dấu nháy đơn và xử lý ký tự đặc biệt
                        escaped_val = str(val).replace("'", "''").replace('\r', '').replace('\n', ' ')
                        set_clauses.append(f"{col} = N'{escaped_val}'")
                set_str = ', '.join(set_clauses)
                # Sử dụng ROW_NUMBER để update theo thứ tự dòng
                update_sql = f";WITH T AS (SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT 1)) AS rn FROM {table_name})\n"
                update_sql += f"UPDATE T SET {set_str} WHERE rn = {idx+1};\n\n"
                f.write(update_sql)
        print(f"Đã tạo file UPDATE: {script_file}")

except pyodbc.Error as e:
    print(f"Lỗi khi kết nối hoặc truy vấn SQL Server: {e}")
finally:
    if 'cursor' in locals() and cursor:
        cursor.close()
    if 'conn' in locals() and conn:
        conn.close()
