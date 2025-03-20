import pandas as pd
import numpy as np
import random
import openpyxl

def generate_pastel_color():
    # 파스텔 색상 생성 (RGB 값이 180-255 사이)
    r = random.randint(180, 255)
    g = random.randint(180, 255)
    b = random.randint(180, 255)
    return f'#{r:02x}{g:02x}{b:02x}'

def find_header_row(file_path):
    # 엑셀 파일의 처음 10행만 읽어옵니다
    df = pd.read_excel(file_path, header=None, nrows=10)
    
    # BKG NO가 포함된 셀의 위치를 찾습니다
    for idx in range(len(df)):
        row_values = df.iloc[idx].astype(str)
        for col_idx, val in enumerate(row_values):
            if 'BKG NO' in str(val):
                print(f"\n=== BKG NO 위치 정보 ===")
                print(f"헤더 행 번호: {idx + 1}")
                print(f"열 번호: {col_idx + 1}")
                print(f"셀 값: {val}")
                return idx
    
    return 0  # 헤더를 찾지 못한 경우 첫 번째 행을 헤더로 사용

# 엑셀 파일 읽기
file_path = "(TIGER) MSC ZOE GT511W - KRPUS.xlsx"
header_row = find_header_row(file_path)

# 헤더 행을 기준으로 데이터 읽기
df = pd.read_excel(file_path, header=header_row)

# 컬럼명 정리 (특수문자 제거 및 공백 처리)
df.columns = df.columns.str.strip()

# 필요한 컬럼들
target_columns = ['BKG NO', 'CNTR NO', 'REMARK', 'PORT']

# BKG NO별 파스텔 색상 매핑 생성
bkg_colors = {}
unique_bkg_nos = df['BKG NO'].unique()
for bkg_no in unique_bkg_nos:
    bkg_colors[bkg_no] = generate_pastel_color()

# 데이터 추출 및 처리
new_data = []
for _, row in df.iterrows():
    bkg_no = row['BKG NO']
    remark = row['REMARK']
    port = row['PORT']
    cntr_nos = str(row['CNTR NO']).split()  # CNTR NO를 공백 기준으로 분리
    
    # CNTR NO가 없는 경우 빈 문자열로 처리
    if pd.isna(cntr_nos[0]):
        cntr_nos = ['']
    
    # 각 CNTR NO에 대해 새로운 행 생성 (길이가 11인 경우만)
    for cntr_no in cntr_nos:
        if len(cntr_no) == 11:  # CNTR NO 길이가 11인 경우만 처리
            new_data.append({
                'BKG NO': bkg_no,
                'CNTR NO': cntr_no,
                'REMARK': remark,
                'PORT': port
            })

# 새로운 데이터프레임 생성
new_df = pd.DataFrame(new_data)

# 새로운 엑셀 파일로 저장
output_file = "processed_data.xlsx"
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    new_df.to_excel(writer, index=False, sheet_name='Data')
    
    # 워크시트 가져오기
    worksheet = writer.sheets['Data']
    
    # 각 행에 대해 BKG NO에 해당하는 색상 적용
    for idx, row in new_df.iterrows():
        # BKG NO 열의 배경색 설정
        cell = worksheet.cell(row=idx+2, column=1)  # +2는 헤더 행과 0-based index 때문
        cell.fill = openpyxl.styles.PatternFill(start_color=bkg_colors[row['BKG NO']][1:],  # # 제거
                                              end_color=bkg_colors[row['BKG NO']][1:],
                                              fill_type='solid')
    
    # 열 너비 자동 조정
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # CNTR NO와 BKG NO 열은 더 넓게 설정
        if column[0].column_letter in ['A', 'B']:  # A열은 BKG NO, B열은 CNTR NO
            adjusted_width = (max_length + 10)  # 여유 공간을 10칸 더 추가
        else:
            adjusted_width = (max_length + 2)  # 다른 열은 기존대로 2칸 여유
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print(f"\n=== 처리 완료 ===")
print(f"원본 데이터 행 수: {len(df)}")
print(f"처리된 데이터 행 수: {len(new_df)}")
print(f"새로운 엑셀 파일이 생성되었습니다: {output_file}")
