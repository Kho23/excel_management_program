import pandas as pd
import re
from datetime import datetime
"""
이 함수는 연월일 의 문자열 데이터를 인자로 받습니다.
DB에 데이터프레임이 저장되기 전에 이 문자열이 날짜로서 유효한지 검사하여
유효하지 않은 경우 날짜확인 이라는 문자열을 붙여서 검사하도록 합니다.
"""
def validate_date(date_str):
    if len(date_str) != 8 :
        return str(date_str)+"(날짜확인)"
    try:
        year = int(date_str[:4])
        month = int(date_str[4:6])
        day = int(date_str[6:])
        datetime(year, month, day)
        return f"{year}-{month}-{day}"
    except ValueError :
        return str(date_str)+"(날짜확인)"

"""
이 함수는 데이터프레임을 인자로 받아서 엑셀 칼럼에 따라 서로 다른 전처리를 해야 합니다.
우선 인자로 받은 데이터프레임이 비어있으면 함수 실행 자체가 불가하니 맨 처음에 비어있는지 검사합니다. 
검사 통과 후 숫자형, 문자형, 날짜형 이 3가지로 구분하여 각 형식에 맞게 전처리를 할 수 있어야 합니다. 
그렇다면 반복문으로 데이터프레임에서 하나씩 가져와서 컬럼명의 형에 따라 분리해서 
"""
def clean_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
    standard = {"num_data":['사번', '번호'], "date_data":['입사일자', '생년월일'], 'str_data':['비고', '부서', '직급', '이름']}
    temp_df = df.copy()
    for col in temp_df.columns:
        temp_df[col] = temp_df[col].fillna("").astype(str).str.strip()
        if col in standard['num_data']:
            temp_df[col] = temp_df[col].astype(str).str.apply(lambda x : re.sub(r'[^0-9]','',x))
        elif col in standard['date_data']:
            temp_df[col] = temp_df[col].astype(str).str.apply(validate_date)
        elif col in standard['str_data']:
            temp_df[col] = temp_df[col].astype(str).str.apply(lambda x : re.sub(r'[^가-힣a-zA-z0-9]','',x))
        temp_df[col] = temp_df[col].astype(str).str.replace(" ", "")
    return temp_df
"""
이 함수는 엑셀 데이터를 DB 에 저장하기 전에 데이터 오염 방지를 위해 데이터를 검수합니다. 
사번이 중복되어 있는 경우 중복된 사번 리스트를 메세지로 표시하고 함수를 종료시킵니다. 
비고에 '삽입필요' 라는 문자열이 있는 경우 행 추가 후 수정되지 않았다는 메세지를 표시하고 함수를 종료시킵니다. 
연락처, 생년월일, 입사일자 에 '확인' 이라는 문자열이 있는 경우 형식이 맞지 않는 데이터가 있다는 메세지를 표시하고 함수를 종료시킵니다. 
"""
def validate_db_data(df, check_col):
    if df is None or df.empty:
        return None
    # check_col = {'UNIQUE_KEY':['사번'], 'REQUIRED_TEXT':['비고', '연락처', '생년월일', '입사일자']}
    error_data = {'status':'success', 'msg':[], 'loc':[]}
    for col in check_col.get('UNIQUE_KEY',[]):
        dup_value = df[col].duplicated(keep='first')
        if dup_value.any():
            dup_data = df.loc[dup_value, col].unique().tolist()
            dup_idx = df[dup_value].index.tolist()
            first_dup_idx = dup_idx[0]
            error_data['status'] = 'error'
            error_data['msg'].append(('사번 중복',f'중복된 사번이 있습니다. 중복된 사번 리스트={dup_data}'))
            error_data['loc'].append(('사번 중복', first_dup_idx))
    for col in check_col.get('REQUIRED_TEXT',[]):
        invalid_date = df[col].astype(str).str.contains('확인')
        if invalid_date.any():
            invalid_data_idx = df[invalid_date].index.tolist()
            first_dup_idx = invalid_data_idx[0]
            error_data['status'] = 'error'
            error_data['msg'].append(('날짜/연락처 오류', f'날짜/연락처 형태가 맞지 않는 데이터가 있습니다. 빨간 칸을 확인해주세요.'))
            error_data['loc'].append(('날짜/연락처 오류', first_dup_idx))
    for col in check_col.get('NEED_INSERT',[]):
        need_insert_value = df[col].astype(str).str.contains('삽입필요')
        if need_insert_value.any():
            need_insert_data_idx = df[need_insert_value].index.tolist()
            first_dup_idx = need_insert_data_idx[0]
            error_data['status'] = 'error'
            error_data['msg'].append(('수정 필요', f'행 추가 후 수정되지 않은 데이터가 있습니다. 비고 란을 확인해주세요.'))
            error_data['loc'].append(('수정 필요',first_dup_idx))
    return error_data


