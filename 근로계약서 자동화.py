import pandas as pd
from datetime import datetime, timedelta
from google.colab import files

# 1️⃣ 파일 업로드
print("📂 인사대장 파일을 업로드하세요.")
uploaded1 = files.upload()
file1_name = list(uploaded1.keys())[0]

print("📂 전자계약서 업로딩 파일을 업로드하세요.")
uploaded2 = files.upload()
file2_name = list(uploaded2.keys())[0]

# 2️⃣ 엑셀 파일 로드
df1 = pd.read_excel(file1_name, sheet_name="Sheet1")             # 기존 계약 데이터
salary_table = pd.read_excel(file1_name, sheet_name="Sheet2")    # 밴드별 급여
location_table = pd.read_excel(file1_name, sheet_name="Sheet3")  # 근무지코드 → 근무장소 테이블
df2 = pd.read_excel(file2_name, sheet_name="Sheet1")             # 신규 입사자 데이터

# 3️⃣ 입사일 기준 필터
df_filtered = df2[df2['입사일'] == '2025-01-01'].copy()

# 4️⃣ 근무지코드와 근무장소 매핑 (Sheet3 기준: 첫 번째 열 = 코드, 두 번째 열 = 장소)
location_map = location_table.set_index(location_table.columns[0])[location_table.columns[1]].to_dict()
df_filtered["근무장소"] = df_filtered["근무지코드"].map(location_map)

# 5️⃣ 신규 데이터 생성
df_new = pd.DataFrame()
df_new["사번"] = df_filtered["사번"]
df_new["구분"] = df_filtered["구분"]
df_new["기준일자"] = "2025-01-01"
df_new["계약시작일"] = "2025-01-01"
df_new["계약종료일"] = (
    datetime.strptime("2025-01-01", "%Y-%m-%d") + timedelta(days=6*30)
).strftime("%Y-%m-%d")
df_new["근무장소"] = df_filtered["근무장소"]
df_new["밴드"] = df_filtered["밴드"]

# 6️⃣ 급여 정보 병합
df_merged = pd.merge(df_new, salary_table, on="밴드", how="left")

# 7️⃣ 기존 컬럼 기준 정리
columns_to_keep = df1.columns.tolist()
df_cleaned = df_merged[columns_to_keep].dropna(axis=1, how='all')

# 8️⃣ 기존 Sheet1과 병합
df_final = pd.concat([df1, df_cleaned], ignore_index=True)

# 9️⃣ 정리
if '밴드' in df_final.columns:
    df_final.drop(columns=['밴드'], inplace=True)
df_final = df_final.dropna(how='all')

# 🔟 저장
final_file = "최종_정리본.xlsx"
df_final.to_excel(final_file, sheet_name="Sheet1", index=False)
print(f"✅ 엑셀 자동화 완료! 저장된 파일명: {final_file}")
files.download(final_file)
