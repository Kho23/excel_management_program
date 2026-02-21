import sys
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5 import uic
from _datetime import date
from datetime import datetime
from PyQt5.QtCore import Qt, pyqtSignal
import sqlite3

ui_data = uic.loadUiType("my_ui.ui")[0]
dialog_data = uic.loadUiType("dialog.ui")[0]

class Dialog(QDialog, dialog_data):
    updated_data = pyqtSignal(object)

    def __init__(self, main_data, file_name):
        super().__init__()
        self.setupUi(self)
        self.conn = sqlite3.connect('db_dept.db')

        self.modified_cell_loc = set()
        self.mod_data = main_data.copy()
        self.file_name = file_name
        self.display_mod_data()

        self.tbl_mod_data.cellChanged.connect(self.change_handler)
        self.btn_create_column.clicked.connect(self.create_column)
        self.tbl_mod_data.customContextMenuRequested.connect(self.show_context)
        self.btn_save_change.clicked.connect(self.save_change)

    def show_context(self, loc):  # 마우스 오른쪽 클릭 시 메뉴 뜨게
        idx = self.tbl_mod_data.indexAt(loc)
        if not idx:
            return
        is_deleted_value = str(self.mod_data.iloc[idx.row()]['is_deleted'])
        menu = QMenu(self)
        if is_deleted_value=='Y':
            btn_undo_delete = menu.addAction("행 삭제 복구")
            btn_undo_delete.triggered.connect(self.undo_delete)
        else:
            btn_delete = menu.addAction("행 삭제")
            btn_delete.triggered.connect(self.delete_data)
        menu.exec_(self.tbl_mod_data.mapToGlobal(loc))

    def delete_data(self):
        selected_row = self.tbl_mod_data.currentRow()
        if selected_row<0:
            return
        target_id = self.mod_data.iloc[selected_row]['사번']
        delete_reason, ok = QInputDialog.getText(self, '삭제 사유 입력', f'사번 {target_id} 데이터의 삭제 사유를 입력하세요.')
        if not ok:
            return
        if not delete_reason:
            QMessageBox.warning(self, '삭제 사유 미입력', '삭제 사유를 입력하지 않았습니다.')
            return
        if ok and delete_reason:
            try:
                if 'is_deleted' in self.mod_data.columns:
                    idx = self.mod_data.columns.get_loc('is_deleted')
                    idx2 = self.mod_data.columns.get_loc('delete_reason')
                    self.mod_data.iat[selected_row, idx] = 'Y'
                    self.mod_data.iat[selected_row, idx2] = str(delete_reason)
            except Exception as e:
                QMessageBox.warning(self, "삭제 실패", f"삭제에 실패했습니다. 에러 내용: {e}")
                return
        self.display_mod_data()
        # self.display_deleted_data()

    def undo_delete(self):
        selected_row = self.tbl_mod_data.currentRow()
        if selected_row<0:
            return
        selected_id = self.mod_data.iloc[selected_row]['사번']
        try:
            self.mod_data.loc[self.mod_data['사번']==selected_id, 'is_deleted'] = 'N'
            self.mod_data.loc[self.mod_data['사번']==selected_id, 'delete_reason'] = '미삭제'
            self.preprocessing()
            self.display_mod_data()
        except Exception as e :
            QMessageBox.warning(self, f'삭제 되돌리기 중 오류 발생, 오류내용: {e}')
            return

    def change_handler(self,row,col):
        self.tbl_mod_data.blockSignals(True) # 무한 루프 방지를 위해 시그널 일시 차단
        tbl_mod_data = self.tbl_mod_data
        data = self.mod_data
        item = self.tbl_mod_data.item(row,col)
        data.iat[row, col] = item.text()
        if '사번' in data.columns:
            emp_id = self.mod_data.iloc[row]['사번']
            co = self.mod_data.columns[col]
            self.modified_cell_loc.add((emp_id, co))
            print(f"변경된 데이터 {self.modified_cell_loc}")
        self.preprocessing()
        self.display_mod_data()
        tbl_mod_data.blockSignals(False) # 수정 완료 후 시그널 차단 해제

    def save_change(self):
        data = self.mod_data
        error_msg = []
        if data is None or data.empty:
            QMessageBox.warning(self, '저장 불가', '저장할 데이터가 존재하지 않습니다.')
            return
        check_col = {
            'unique_key':['사번'],
            'need_modify':['비고'],
            'need_check':['연락처', '입사일자', '생년월일'],
        }
        error_loc = None
        for col in check_col['unique_key']:
            duplicated_val = data[col].duplicated()
            if duplicated_val.any():
                if error_loc is None:
                    error_row = data.index[duplicated_val][0]
                    error_col = data.columns.get_loc(col)
                    error_loc = (error_row, error_col)
                duplicated_data = data.loc[duplicated_val, col].unique().tolist()
                error_msg.append(f'중복된 사번이 있습니다. 중복된 사번 목록:{duplicated_data}')
        for col in check_col['need_modify']:
            need_modify_val = data[col].astype(str).str.contains('삽입필요', na=False)
            if need_modify_val.any():
                if error_loc is None:
                    error_row = data.index[need_modify_val][0]
                    error_col = data.columns.get_loc(col)
                    error_loc = (error_row, error_col)
                error_msg.append("행 추가 후 수정되지 않은 데이터가 있습니다. 비고 란에 '삽입 필요' 가 있는 행을 수정 후 저장해주세요.")
        for col in check_col['need_check']:
            need_modify_val = data[col].astype(str).str.contains('확인', na=False)
            if need_modify_val.any():
                if error_loc is None:
                    error_row = data.index[need_modify_val][0]
                    error_col = data.columns.get_loc(col)
                    error_loc = (error_row, error_col)
                error_msg.append('날짜&연락처 형식에 맞지 않는 데이터가 있습니다. 빨간 칸을 확인하여 수정 후 저장해주세요.')
        if error_msg:
            full_error_msg = '\n'.join(error_msg)
            QMessageBox.warning(self, '데이터 오류 확인 필요',  f'저장 중 데이터 형식 오류로 에러가 발생했습니다.\n오류 내용: {full_error_msg}')
            if error_loc:
                row, col = error_loc
                self.tbl_mod_data.setCurrentCell(row, col)
                self.tbl_mod_data.setFocus()
            return
        try:
            self.updated_data.emit(data)
            self.modified_cell_loc.clear()
            QMessageBox.information(self, '저장 완료', '변경 내역이 저장되었습니다.')
            self.accept()
        except Exception as e :
            QMessageBox.warning(self, "저장 중 오류", f"데이터 저장 중 오류가 발생했습니다.\n오류내용: {e}")
            return

    def preprocessing(self):
        data = self.mod_data

        def validate_date(date_str):
            if len(date_str) > 8 :
                date_str = date_str[:8]
            if len(date_str) < 8:
                return f"{date_str}(길이확인)"
            try:
                year = int(date_str[:4])
                month = int(date_str[4:6])
                day = int(date_str[6:])
                datetime(year, month, day) # 날짜 데이터로 만들어보고 에러가 터지지 않으면 반환
                return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            except ValueError:
                return f"{date_str}(날짜확인)"

        if data is not None:
            for col in data.columns:
                data[col] = data[col].fillna("").astype(str).str.strip() # 공백 모두 제거하고
                if col == "번호":
                    data[col] = range(1,len(data)+1) # 번호 자동 정렬 시키기
                elif col in ["입사일자", "생년월일", "연락처", "사번"]: # 숫자가 들어가는 데이터
                    data[col] = data[col].astype(str).str.replace(r'([^0-9])', "", regex=True).str[:]
                    if col == '연락처':
                        data[col] = data[col].apply(lambda x: f"{x[:3]}-{x[3:7]}-{x[7:]}" if 10 <= len(x) <= 11 else x+"(길이확인)")
                    if col in ['생년월일','입사일자']:
                        data[col] = data[col].apply(validate_date)
                else:
                    data[col] = data[col].astype(str).str.replace(r'[^가-힣a-zA-Z0-9\s]', "", regex=True)
                data[col] = data[col].astype(str).str.replace("nan", "")
                data[col] = data[col].astype(str).str.replace(" ", "")

    def display_mod_data(self):
        data = self.mod_data
        data.sort_values(by='is_deleted', ascending=True, inplace=True)
        data.reset_index(drop=True, inplace=True)
        data['번호'] = range(1, len(data)+1)
        self.tbl_mod_data.blockSignals(True)
        rows, cols = data.shape
        self.tbl_mod_data.setRowCount(rows)
        self.tbl_mod_data.setColumnCount(cols)
        self.tbl_mod_data.setHorizontalHeaderLabels(data.columns)
        for i in range(rows):
            current_id = self.mod_data.iloc[i]["사번"]
            for j in range(cols):
                value = str(data.iloc[i, j])
                item = QTableWidgetItem(value)
                if str(data.iloc[i]['is_deleted']) == 'Y':
                    item.setBackground(Qt.lightGray)
                if '확인' in value:
                    item.setBackground(Qt.red)
                elif (str(current_id), self.mod_data.columns[j]) in self.modified_cell_loc:
                    item.setBackground(Qt.yellow)
                item.setTextAlignment(Qt.AlignCenter)
                self.tbl_mod_data.setItem(i, j, item)
        for col in ['file_name', 'is_deleted', 'delete_reason']:
            col_idx = self.mod_data.columns.get_loc(col)
            self.tbl_mod_data.setColumnHidden(col_idx, True)
        self.tbl_mod_data.blockSignals(False)

    def create_column(self):
        if self.mod_data is not None :
            new_row = pd.DataFrame([{
                "번호":"0",
                "이름": "신규",
                "직급":"신규",
                "부서": "미정",
                "사번":"00000000",
                "입사일자": "2026-01-01",
                "생년월일":"2000-01-01",
                "연락처":"010-0000-0000",
                "비고":"삽입 필요"
            }])
            try:
                self.mod_data = pd.concat([self.mod_data, new_row], ignore_index=True)
                self.preprocessing()
                self.display_mod_data()
            except Exception as e :
                QMessageBox.warning(self, "행 추가 실패", f"행 추가에 실패했습니다. 에러 내용:{e}")
                return
        else:
            QMessageBox.warning(self, '엑셀 데이터 없음', '엑셀 데이터 추가 후 사용해주세요.')
            return

class Main(QMainWindow, ui_data):
    data = pd.DataFrame()
    stat_data = pd.DataFrame()

    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.conn = sqlite3.connect('db_dept.db')
        self.file_name='엑셀 파일.xlsx'

        self.btn_load_excel.clicked.connect(self.load_excel)
        self.btn_modify.clicked.connect(self.open_dialog)
        self.btn_save.clicked.connect(self.save_excel)
        self.btn_save_db.clicked.connect(self.save_db)
        self.btn_load_db.clicked.connect(self.load_db)
        self.check_db()
        self.load_deleted_data()

    def open_dialog(self):
        if self.data.empty:
            QMessageBox.warning(self, '데이터 없음', '수정할 데이터가 없습니다.')
            return
        dlg = Dialog(self.data, self.file_name)
        dlg.updated_data.connect(self.apply_dlg_changes)
        dlg.exec_()

    def apply_dlg_changes(self, update_data):
        self.data = update_data.copy()
        self.preprocessing()
        self.display_data()
        self.display_deleted_data()
        self.display_stat()

    def fetch_db(self, file_name):
        cursor = self.conn.cursor()
        try:
            cursor.execute('SELECT COUNT(*) FROM dept_data WHERE file_name=?', (file_name,))
            if cursor.fetchone()[0] == 0:
                return None
            df = pd.read_sql('SELECT * FROM dept_data WHERE file_name = ? ', self.conn, params=(file_name, ))
            return df
        except Exception as e:
            print(f"엑셀 DB 조회 중 에러 발생 에러내용:{e}")
            QMessageBox.warning(self, '엑셀 조회 중 오류 발생', f'엑셀 조회 중 오류가 발생했습니다. 오류 내용: {e}')
            return None

    def check_db(self):
        cursor = self.conn.cursor()
        cursor.execute(
            'CREATE TABLE IF NOT EXISTS file_log(file_name TEXT PRIMARY KEY, upload_date TIMESTAMP)')
        cursor.execute(
            'CREATE TABLE IF NOT EXISTS dept_data('
            '번호 TEXT, 이름 TEXT, 직급 TEXT, 부서 TEXT, 사번 TEXT, 입사일자 TEXT, 생년월일 TEXT, 연락처 TEXT, 비고 TEXT, file_name TEXT, is_deleted TEXT DEFAULT N, delete_reason TEXT DEFAULT 미삭제)',
        )
        cursor.execute("SELECT file_name FROM file_log")
        db_table_list = cursor.fetchall()
        print(f"현재 db에 저장된 엑셀 목록: {db_table_list}")
        self.cmb_db_list.clear()
        if db_table_list is not None:
            for db in db_table_list:
                table_name = db[0].split('/')[-1]
                self.cmb_db_list.addItem(table_name, db)
        else : return

    def save_db(self):
        if self.data['사번'].duplicated().any():
            duplicated_id = self.data[self.data['사번'].duplicated()]['번호']
            QMessageBox.warning(self, '중복된 사번 발견', f"사번이 중복된 데이터가 있습니다.\n중복된 데이터 번호는 {list(duplicated_id)} 입니다.\n수정 후 'DB에 저장' 버튼을 눌러주세요.")
            self.tbl_data.setCurrentCell(int(list(duplicated_id)[0])-1, 4)
            self.tbl_data.setFocus()
            return
        elif self.data['비고'].str.contains('삽입필요',na=False).any():
            added_col_list = self.data[self.data['비고'].str.contains('삽입필요', na=False)].index.to_list()[0]
            QMessageBox.warning(self, '수정 완료되지 않은 데이터 존재', f"신규 행 추가 후 {added_col_list+1}번 데이터를 수정하지 않았습니다.\n수정 및 비고의 '삽입 필요' 제거 후 'DB에 저장' 버튼을 눌러주세요.")
            self.tbl_data.setCurrentCell(added_col_list, 8)
            self.tbl_data.setFocus()
            return
        elif self.data[['입사일자', '연락처', '생년월일']].apply(lambda x : x.str.contains('확인')).any().any():
            QMessageBox.warning(self, "날짜 & 연락처 데이터 확인 필요", "형식에 맞지 않는 데이터가 있습니다. 빨간 칸을 확인 후 'DB에 저장' 버튼을 눌러주세요.")
            return
        else:
            cursor = self.conn.cursor()
            save_df = self.data.copy()
            save_df['file_name'] = self.file_name
            for _, row in save_df.iterrows():
                cursor.execute('SELECT COUNT(*) FROM dept_data where file_name=? AND 사번=?', (self.file_name,row['사번']))
                if  cursor.fetchone()[0] > 0 :
                    sql = "UPDATE dept_data SET 번호=?, 이름=?, 직급=?, 부서=?, 사번=?, 입사일자=?, 생년월일=?, 연락처=?, 비고=?, file_name=?, is_deleted = ?, delete_reason=? WHERE file_name =? AND 사번 = ?"
                    cursor.execute(sql,
                           (row['번호'], row['이름'], row['직급'], row['부서'], row['사번'], row['입사일자'], row['생년월일'],
                            row['연락처'], row['비고'], self.file_name,row['is_deleted'], row['delete_reason'],  self.file_name, row['사번']))
                else :
                    sql = "INSERT INTO dept_data (번호, 이름, 직급, 부서, 사번, 입사일자, 생년월일, 연락처, 비고, file_name, is_deleted, delete_reason) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
                    cursor.execute(sql,
                           (row['번호'], row['이름'], row['직급'], row['부서'], row['사번'], row['입사일자'], row['생년월일'],
                            row['연락처'], row['비고'], self.file_name,row["is_deleted"], row['delete_reason']))
            cursor.execute('SELECT count(*) FROM file_log WHERE file_name = ?', (self.file_name,))
            if cursor.fetchone()[0] == 0 :
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cursor.execute('INSERT INTO file_log values(?,?)', (self.file_name, now))
            self.conn.commit()
            QMessageBox.information(self, "DB 저장 완료", "DB 저장이 완료되었습니다.")
            self.check_db()
            self.display_data()

    def load_db(self):
        selected_db = self.cmb_db_list.currentData()[0]
        print(f"선택된 디비 {selected_db}")
        if not selected_db:
            return
        db_data = self.fetch_db(selected_db)
        if db_data is not None:
            self.data = db_data
            self.file_name = selected_db
            self.preprocessing()
            self.display_data()
            self.display_deleted_data()
            self.display_stat()
            QMessageBox.information(self, 'DB 조회 완료', 'DB에서 데이터를 가져왔습니다.')
        else:
            QMessageBox.warning(self, 'DB 조회 실패', 'DB에서 데이터를 가져오는데 실패했습니다.')

    def load_excel(self):
        fname = QFileDialog.getOpenFileName(self, '조회할 엑셀 파일 선택', '조회할 엑셀 파일을 선택해주세요.', 'Excel Files (*.xlsx)')
        if not fname[0]:
            return
        self.file_name = fname[0]
        db_data = self.fetch_db(self.file_name)
        if db_data is not None:
            print("저장된 엑셀을 불러옵니다.")
            self.data = db_data
        else :
            print("새로운 엑셀 파일입니다.")
            self.data = pd.read_excel(fname[0], sheet_name="인사정보", header=2, dtype=str)
            if 'is_deleted' not in self.data.columns:
                self.data['is_deleted'] = 'N'
            if 'deleted_reason' not in self.data.columns:
                self.data['delete_reason'] = '삭제 안됨'
            if 'file_name' not in self.data.columns:
                self.data['file_name'] = fname[0]
            self.preprocessing()
            self.save_db()
            self.conn.commit()
        self.display_data()
        self.display_deleted_data()
        self.display_stat()
        self.check_db()

    def load_deleted_data(self):
        cursor = self.conn.cursor()
        try:
            cursor.execute('SELECT COUNT(*) FROM dept_data WHERE is_deleted = ? AND file_name = ?', ('Y',self.file_name)) # 삭제됨 처리된 데이터만 가져와서
            if cursor.fetchone()[0] == 0:
                return None
            df = pd.read_sql("SELECT * FROM dept_data WHERE is_deleted = ? AND file_name=?", self.conn, params=("Y", self.file_name))
            return df
        except Exception as e :
            QMessageBox.warning(self, '삭제 데이터 조회 중 오류 발생', f"삭제 데이터 조회 중 오류가 발생했습니다.\n오류내용:{e}")
            return None

    def display_deleted_data(self):
        self.tbl_deleted_data.blockSignals(True)
        deleted_data = self.load_deleted_data()
        deleted_table = self.tbl_deleted_data
        if deleted_data is not None:
            deleted_data.rename(columns={'delete_reason': '삭제 사유'}, inplace=True)
            deleted_data.drop('file_name', axis=1, inplace=True)
            deleted_data.drop('is_deleted', axis=1, inplace=True)
            deleted_table.setRowCount(len(deleted_data))
            deleted_table.setColumnCount(len(deleted_data.columns))
            deleted_table.setHorizontalHeaderLabels(deleted_data.columns)
            for i in range(len(deleted_data)):
                for j in range(len(deleted_data.columns)):
                    val = str(deleted_data.iloc[i, j])
                    item = QTableWidgetItem(val)
                    item.setTextAlignment(Qt.AlignCenter)
                    self.tbl_deleted_data.setItem(i, j, item)
        self.tbl_deleted_data.blockSignals(False)

    def preprocessing(self):
        data = self.data

        def validate_date(date_str):
            if len(date_str) > 8 :
                date_str = date_str[:8]
            if len(date_str) < 8:
                return f"{date_str}(길이확인)"
            try:
                year = int(date_str[:4])
                month = int(date_str[4:6])
                day = int(date_str[6:])
                datetime(year, month, day)
                return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"
            except ValueError:
                return f"{date_str}(날짜확인)"

        if data is not None:
            for col in data.columns:
                data[col] = data[col].fillna("").astype(str).str.strip() # 공백 모두 제거하고
                if col == "번호":
                    data[col] = range(1,len(data)+1) # 번호 자동 정렬 시키기
                elif col in ["입사일자", "생년월일", "연락처", "사번"]: # 숫자가 들어가는 데이터
                    data[col] = data[col].astype(str).str.replace(r'([^0-9])', "", regex=True).str[:]
                    if col == '연락처':
                        data[col] = data[col].apply(lambda x: f"{x[:3]}-{x[3:7]}-{x[7:]}" if 10 <= len(x) <= 11 else x+"(길이확인)")
                    if col in ['생년월일','입사일자']:
                        data[col] = data[col].apply(validate_date)
                else:
                    data[col] = data[col].astype(str).str.replace(r'[^가-힣a-zA-Z0-9\s]', "", regex=True)
                data[col] = data[col].astype(str).str.replace("nan", "")
                data[col] = data[col].astype(str).str.replace(" ", "")

    def display_data(self):
        self.tbl_data.blockSignals(True)

        if 'is_deleted' in self.data.columns:
            self.data.sort_values(by=['is_deleted'], ascending=True, inplace=True)
            self.data.reset_index(drop=True, inplace=True)
        display_data = self.data[self.data['is_deleted']=='N'].copy()
        table = self.tbl_data  # 테이블 위젯에 엑셀 데이터 띄우기
        table.setRowCount(len(display_data))
        table.setColumnCount(len(display_data.columns))
        table.setHorizontalHeaderLabels(display_data.columns)

        for i in range(len(display_data)):
                for j in range(len(display_data.columns)):
                    if j == list(display_data.columns).index("번호"):
                        val = str(i+1)
                    else:
                        val = str(display_data.iloc[i, j])
                    item = QTableWidgetItem(val)
                    if "확인" in val:
                        item.setBackground(Qt.red)
                    item.setTextAlignment(Qt.AlignCenter)
                    table.setItem(i, j, item)

        for col_name in ['is_deleted', 'file_name', 'delete_reason']:
            if col_name in display_data.columns:
                col_idx = display_data.columns.get_loc(col_name)
                table.setColumnHidden(col_idx, True)

        self.stat_data = display_data[display_data['is_deleted']== 'N'].copy()
        self.display_stat()
        self.tbl_data.blockSignals(False)

    def display_stat(self):
        stat_data = self.stat_data
        if stat_data.empty:
            self.ln_total_emp.setText("0")
            self.ln_emp_age_avg.setText("0")
            self.ln_total_longterm_emp.setText("0")
            self.ln_oldest_emp.setText("-")
            self.tbl_dept_stat.setRowCount(0)
            self.tbl_long_term.setRowCount(0)
            return
        today_pd = pd.to_datetime(date.today())
        hired_date = pd.to_datetime(stat_data['입사일자'], format='%Y-%m-%d', errors='coerce')
        birth_date = pd.to_datetime(stat_data['생년월일'], format='%Y-%m-%d', errors='coerce')

        stat_data['연령'] = ((today_pd - birth_date).dt.days / 365.25).fillna(0).astype(int)
        stat_data['근속연수'] = ((today_pd - hired_date).dt.days / 365.25).fillna(0).astype(int)

        age_avg = stat_data['연령'].mean()
        long_term_emp = len(stat_data[stat_data['근속연수'] > 10])
        oldest_name = stat_data.loc[stat_data['연령'].idxmax(), '이름']

        dept_stat_table = self.tbl_dept_stat
        long_term_table = self.tbl_long_term
        dept_stats = stat_data.groupby('부서').agg({
            '사번': 'size',
            '연령': ['max', 'mean'],
            '근속연수': lambda x: (x > 10).sum()
        }).reset_index()

        dept_stat_header = ['부서명', '인원', '최고령자', '평균 연령', '장기근속자 수']

        dept_stat_table.setColumnCount(len(dept_stat_header))
        dept_stat_table.setRowCount(len(dept_stats))
        dept_stat_table.setHorizontalHeaderLabels(dept_stat_header)

        for i in range(len(dept_stats)):
            for j in range(len(dept_stats.columns)):
                val = dept_stats.iloc[i, j]
                if j == 3:
                    item = QTableWidgetItem(str(round(val, 1)))
                else:
                    item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                dept_stat_table.setItem(i, j, item)

        long_term_data = stat_data[stat_data['근속연수'] > 10]
        long_term_column = ['이름', '부서', '사번', '입사일자', '근속연수']
        long_term_stats = long_term_data[long_term_column]

        long_term_table.setColumnCount(len(long_term_column))
        long_term_table.setRowCount(len(long_term_stats))
        long_term_table.setHorizontalHeaderLabels(long_term_column)

        for i in range(len(long_term_stats)):
            for j in range(len(long_term_column)):
                val = long_term_stats.iloc[i, j]
                if j == 4:
                    item = QTableWidgetItem(str(val) + "년")
                else:
                    item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                long_term_table.setItem(i, j, item)

        self.ln_total_emp.setText(str(len(stat_data)))
        self.ln_ref_date.setText(str(date.today()))
        self.ln_emp_age_avg.setText(str(round(age_avg, 1)))
        self.ln_total_longterm_emp.setText(str(long_term_emp))
        self.ln_oldest_emp.setText(str(oldest_name))

    def save_excel(self):
        # 엑셀 저장할때 원하는 열만 나오도록 수정하기
        self.preprocessing()
        save_data = self.data.copy()
        save_data = save_data[save_data['is_deleted']!='Y']
        today = str(date.today()).replace("-", "")
        try:
            col_num = min(len(save_data.columns), 9)
            save_data = save_data.iloc[:, :col_num]
            save_data.to_excel(f'회사 사원 정보_{today}.xlsx', index=False)
            QMessageBox.information(self, "엑셀 저장 완료", "엑셀 파일이 저장되었습니다." )
        except Exception as e:
            QMessageBox.warning(self, "엑셀 저장 오류", f"엑셀 저장중 에러가 발생했습니다. 에러 내용: {e}")
            return

    # def change_handler(self,row,col):
    #     self.tbl_data.blockSignals(True) # 무한 루프 방지를 위해 시그널 일시 차단
    #     tbl_data = self.tbl_data
    #     data = self.data
    #     item = self.tbl_data.item(row,col)
    #     data.iat[row, col] = item.text()
    #     if '사번' in data.columns:
    #         emp_id = self.data.iloc[row]['사번']
    #         co = self.data.columns[col]
    #         self.modified_cell_loc.add((emp_id, co))
    #         print(f"변경된 데이터 {self.modified_cell_loc}")
    #     self.preprocessing()
    #     self.display_data()
    #     self.display_stat()
    #     tbl_data.blockSignals(False) # 수정 완료 후 시그널 차단 해제

if __name__=="__main__":
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())