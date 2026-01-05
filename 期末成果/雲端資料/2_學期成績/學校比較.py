import pandas as pd
from sqlalchemy import create_engine, text
from openpyxl import load_workbook
import os

# PostgreSQL 連線設定
conn_str = "postgresql+psycopg2://postgres:eric1611@35.229.163.253:5432/postgres"
engine = create_engine(conn_str)

# 從 document schema 抓取資料表 "學期成績"
df = pd.read_sql('SELECT * FROM document."學期成績"', con=engine)

# 清洗數據欄位（轉數字）
subjects = ['使用者在國文測驗的分數', '使用者在數學測驗的分數', '使用者在英文測驗的分數']
for subject in subjects:
    df[subject] = pd.to_numeric(df[subject], errors='coerce')

# 及格標準
pass_score = 60

# 按學校代碼分組
grouped = df.groupby('學校代碼')

# 建立空結果表
result = pd.DataFrame()

# 各科分析
for subject in subjects:
    tmp = grouped[subject].agg([
        ('平均分數', 'mean'),
        ('最高分', 'max'),
        ('最低分', 'min'),
        ('標準差', 'std'),
        ('學生數', 'count'),
        ('及格人數', lambda x: (x >= pass_score).sum()),
        ('及格率', lambda x: (x >= pass_score).mean())
    ])
    tmp = tmp.add_prefix(subject + '_')
    result = tmp if result.empty else pd.concat([result, tmp], axis=1)

result = result.reset_index()

# 及格率轉百分比
for subject in subjects:
    result[subject + '_及格率'] = result[subject + '_及格率'] * 100

# 欄位名稱簡化
rename_dict = {
    '學校代碼': '學校代碼',
    '使用者在國文測驗的分數_平均分數': '國文測驗平均',
    '使用者在國文測驗的分數_最高分': '國文測驗最高分',
    '使用者在國文測驗的分數_最低分': '國文測驗最低分',
    '使用者在國文測驗的分數_標準差': '國文測驗標準差',
    '使用者在國文測驗的分數_學生數': '國文測驗學生數',
    '使用者在國文測驗的分數_及格人數': '國文測驗及格人數',
    '使用者在國文測驗的分數_及格率': '國文測驗及格率',
    '使用者在數學測驗的分數_平均分數': '數學測驗平均',
    '使用者在數學測驗的分數_最高分': '數學測驗最高分',
    '使用者在數學測驗的分數_最低分': '數學測驗最低分',
    '使用者在數學測驗的分數_標準差': '數學測驗標準差',
    '使用者在數學測驗的分數_學生數': '數學測驗學生數',
    '使用者在數學測驗的分數_及格人數': '數學測驗及格人數',
    '使用者在數學測驗的分數_及格率': '數學測驗及格率',
    '使用者在英文測驗的分數_平均分數': '英文測驗平均',
    '使用者在英文測驗的分數_最高分': '英文測驗最高分',
    '使用者在英文測驗的分數_最低分': '英文測驗最低分',
    '使用者在英文測驗的分數_標準差': '英文測驗標準差',
    '使用者在英文測驗的分數_學生數': '英文測驗學生數',
    '使用者在英文測驗的分數_及格人數': '英文測驗及格人數',
    '使用者在英文測驗的分數_及格率': '英文測驗及格率'
}
result = result.rename(columns=rename_dict)

# 匯出 Excel
excel_path = '學校分析結果.xlsx'
result.to_excel(excel_path, index=False)

# 自動調整欄寬
wb = load_workbook(excel_path)
ws = wb.active
for column_cells in ws.columns:
    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length * 2
wb.save(excel_path)

print("✅ 已從資料庫分析並輸出：學校分析結果.xlsx（欄寬已調整）")

# ✅ 將結果存入 result schema 中的資料表
table_name = os.path.splitext(os.path.basename(excel_path))[0]  # "學校分析結果"
with engine.connect() as conn:
    conn.execute(text("CREATE SCHEMA IF NOT EXISTS result;"))  # 確保 schema 存在

# 寫入資料庫（覆蓋模式）
result.to_sql(name=table_name, con=engine, schema='result', if_exists='replace', index=False)
print(f"📤 分析結果已匯入 result.\"{table_name}\" 資料表")
