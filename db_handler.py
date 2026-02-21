import pandas as pd
from utils import validate_db_data
def fetch_db(file_name, conn):
    if file_name is None:
        return pd.DataFrame()
    try:
        df = pd.read_sql('SELECT * FROM dept_data WHERE file_name=?', conn, params=(file_name,))
        if not df.empty:
            return df
    except Exception as e:
        print(f"db 조회 중 에러 발생 에러내용={e}")
        return pd.DataFrame()
    return pd.DataFrame()

def save_db(df, conn, file_name):
    if df is None or df.empty:
        return None
    cur = conn.cursor()
    check_col = {
        'UNIQUE_KEY':['사번'],
        'REQUIRED_TEXT':['연락처', '생년월일', '입사일자'],
        'NEED_INSERT':['비고']
    }
    temp_df = df.copy()
    temp_df['file_name'] = file_name
    error_data = validate_db_data(temp_df, check_col)
    if error_data['status'] == 'success':
        cur.execute("", )





