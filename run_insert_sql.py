import pyodbc

def read_connect_string(filename):
    """
    Đọc chuỗi kết nối từ file.
    """
    with open(filename, 'r', encoding='utf-8') as f:
        return f.read().strip()

def run_sql_file(connect_string, sql_file):
    """
    Thực thi tất cả các câu lệnh SQL trong file sql_file lên DB theo connect_string.
    """
    with open(sql_file, 'r', encoding='utf-8') as f:
        sql_content = f.read()
    statements = [stmt.strip() for stmt in sql_content.split(';') if stmt.strip()]
    conn = pyodbc.connect(connect_string)
    cursor = conn.cursor()
    for stmt in statements:
        try:
            cursor.execute(stmt)
        except Exception as e:
            print(f"Error executing: {stmt}\n{e}")
    conn.commit()
    cursor.close()
    conn.close()
    print("All SQL statements executed.")