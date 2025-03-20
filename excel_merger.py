import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import numpy as np
import random
import openpyxl
from pathlib import Path
import os

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Merger")
        self.root.geometry("800x600")
        
        # 파일 리스트를 저장할 변수
        self.file_list = []
        
        # GUI 구성요소 생성
        self.create_widgets()
        
        # 드래그 앤 드롭 설정
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.handle_drop)

    def create_widgets(self):
        # 메인 프레임을 좌우로 분할
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 왼쪽 프레임 (파일 목록)
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # 오른쪽 프레임 (Summary)
        right_frame = ttk.LabelFrame(main_frame, text="Summary", padding="5")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # 파일 리스트 프레임
        list_frame = ttk.LabelFrame(left_frame, text="파일 목록", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 파일 리스트박스
        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        # 버튼 프레임
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 파일 추가 버튼
        add_button = ttk.Button(button_frame, text="파일 추가", command=self.add_files)
        add_button.pack(side=tk.LEFT, padx=5)
        
        # 파일 제거 버튼
        remove_button = ttk.Button(button_frame, text="선택 제거", command=self.remove_files)
        remove_button.pack(side=tk.LEFT, padx=5)
        
        # 처리 시작 버튼
        process_button = ttk.Button(button_frame, text="처리 시작", command=self.process_files)
        process_button.pack(side=tk.RIGHT, padx=5)
        
        # Summary 정보 표시 레이블들
        self.total_bkg_label = ttk.Label(right_frame, text="총 BKG NO 수: 0")
        self.total_bkg_label.pack(fill=tk.X, pady=2)
        
        self.total_cntr_label = ttk.Label(right_frame, text="총 CNTR NO 수: 0")
        self.total_cntr_label.pack(fill=tk.X, pady=2)
        
        self.total_files_label = ttk.Label(right_frame, text="처리된 파일 수: 0")
        self.total_files_label.pack(fill=tk.X, pady=2)
        
        # 상태 표시 레이블
        self.status_label = ttk.Label(left_frame, text="파일을 드래그 앤 드롭하거나 '파일 추가' 버튼을 클릭하세요")
        self.status_label.pack(fill=tk.X, pady=5)

    def handle_drop(self, event):
        # 드래그 앤 드롭으로 받은 파일 경로 처리
        files = event.data.split()
        for file in files:
            # 중괄호와 따옴표 제거
            file = file.strip('{}"')
            
            # 파일 경로에서 파일명 추출
            file_name = os.path.basename(file)
            
            # 파일명에서 확장자만 분리하여 확인
            name, ext = os.path.splitext(file_name)
            if ext.lower() == '.xlsx':
                self.file_list.append(file)
                self.listbox.insert(tk.END, file_name)
            else:
                messagebox.showwarning("경고", f"엑셀 파일만 처리 가능합니다: {file_name}")
        self.update_status()

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xls;*.xlsx")]
        )
        for file in files:
            self.file_list.append(file)
            self.listbox.insert(tk.END, os.path.basename(file))
        self.update_status()

    def remove_files(self):
        selected = self.listbox.curselection()
        for index in reversed(selected):
            self.listbox.delete(index)
            self.file_list.pop(index)
        self.update_status()

    def update_status(self):
        self.status_label.config(text=f"총 {len(self.file_list)}개의 파일이 선택되었습니다")

    def find_header_row(self, file_path):
        df = pd.read_excel(file_path, header=None, nrows=10)
        for idx in range(len(df)):
            row_values = df.iloc[idx].astype(str)
            for col_idx, val in enumerate(row_values):
                if 'BKG NO' in str(val):
                    return idx
        return 0

    def find_remark_columns(self, df):
        # REMARK 열들 찾기 (대소문자 구분 없이)
        remark_cols = {}
        for col in df.columns:
            col_str = str(col).upper()
            if 'REMARK(CS)' in col_str:
                remark_cols['cs'] = col
            elif 'REMARK' in col_str:
                remark_cols['normal'] = col
        return remark_cols

    def update_summary(self, total_bkg, total_cntr, total_files):
        self.total_bkg_label.config(text=f"총 BKG NO 수: {total_bkg}")
        self.total_cntr_label.config(text=f"총 CNTR NO 수: {total_cntr}")
        self.total_files_label.config(text=f"처리된 파일 수: {total_files}")

    def process_files(self):
        if not self.file_list:
            messagebox.showwarning("경고", "처리할 파일을 선택해주세요")
            return

        try:
            all_data = []
            total_bkg = 0
            total_cntr = 0
            all_bkg_nos = set()  # 모든 BKG NO를 저장할 set
            
            # 먼저 모든 파일에서 BKG NO를 수집
            for file_path in self.file_list:
                header_row = self.find_header_row(file_path)
                df = pd.read_excel(file_path, header=header_row)
                df.columns = df.columns.str.strip()
                all_bkg_nos.update(df['BKG NO'].unique())
            
            # BKG NO별 파스텔 색상 매핑 생성 (한 번만)
            bkg_colors = {}
            for bkg_no in all_bkg_nos:
                bkg_colors[bkg_no] = self.generate_pastel_color()
            
            # 데이터 처리
            for file_path in self.file_list:
                header_row = self.find_header_row(file_path)
                df = pd.read_excel(file_path, header=header_row)
                df.columns = df.columns.str.strip()
                
                # REMARK 열들 찾기
                remark_cols = self.find_remark_columns(df)
                if not remark_cols:
                    print(f"경고: {file_path}에서 REMARK 열을 찾을 수 없습니다.")
                    continue
                
                # 데이터 추출 및 처리
                for _, row in df.iterrows():
                    bkg_no = row['BKG NO']
                    remark = row[remark_cols.get('normal', '')] if remark_cols.get('normal') and pd.notna(row[remark_cols['normal']]) else ''
                    remark_cs = row[remark_cols.get('cs', '')] if remark_cols.get('cs') and pd.notna(row[remark_cols['cs']]) else ''
                    port = row['PORT'] if pd.notna(row['PORT']) else ''
                    cntr_nos = str(row['CNTR NO']).split() if pd.notna(row['CNTR NO']) else ['']
                    
                    if pd.isna(cntr_nos[0]):
                        cntr_nos = ['']
                    
                    for cntr_no in cntr_nos:
                        if len(cntr_no) == 11:
                            all_data.append({
                                'BKG NO': bkg_no,
                                'CNTR NO': cntr_no,
                                'REMARK': remark,
                                'REMARK(CS)': remark_cs,
                                'PORT': port
                            })
                            total_cntr += 1
                
                total_bkg += len(df['BKG NO'].unique())

            # 새로운 데이터프레임 생성
            new_df = pd.DataFrame(all_data)
            
            # 새로운 엑셀 파일로 저장
            output_file = "merged_data.xlsx"
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 데이터 시트
                new_df.to_excel(writer, index=False, sheet_name='Data')
                worksheet = writer.sheets['Data']
                
                # BKG NO별 색상 적용
                for idx, row in new_df.iterrows():
                    cell = worksheet.cell(row=idx+2, column=1)
                    cell.fill = openpyxl.styles.PatternFill(
                        start_color=bkg_colors[row['BKG NO']][1:],
                        end_color=bkg_colors[row['BKG NO']][1:],
                        fill_type='solid'
                    )
                
                # 열 너비 조정
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    if column[0].column_letter in ['A', 'B']:
                        adjusted_width = (max_length + 10)
                    else:
                        adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

            # Summary 정보 업데이트
            self.update_summary(total_bkg, total_cntr, len(self.file_list))
            
            messagebox.showinfo("완료", f"처리가 완료되었습니다.\n\n결과 파일: {output_file}")

        except Exception as e:
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다:\n{str(e)}")

    def generate_pastel_color(self):
        r = random.randint(180, 255)
        g = random.randint(180, 255)
        b = random.randint(180, 255)
        return f'#{r:02x}{g:02x}{b:02x}'

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelMergerApp(root)
    root.mainloop() 