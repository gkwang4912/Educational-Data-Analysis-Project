import pandas as pd
import pandas as pd
from openpyxl import load_workbook

# 讀入 Excel 資料
df = pd.read_excel('學期成績.xlsx')

# 欄位如果是文字，先轉成數字型態（避免有空白或異常資料影響）
for subject in ['使用者在國文測驗的分數', '使用者在數學測驗的分數', '使用者在英文測驗的分數']:
    df[subject] = pd.to_numeric(df[subject], errors='coerce')

# 及格標準
pass_score = 60

# 分組計算
grouped = df.groupby('學校代碼')

result = pd.DataFrame()

for subject in ['使用者在國文測驗的分數', '使用者在數學測驗的分數', '使用者在英文測驗的分數']:
    tmp = grouped[subject].agg([
        ('平均分數', 'mean'),
        ('最高分', 'max'),
        ('最低分', 'min'),
        ('標準差', 'std'),
        ('學生數', 'count'),
        ('及格人數', lambda x: (x >= pass_score).sum()),
        ('及格率', lambda x: (x >= pass_score).mean())
    ])
    # 欄位名稱加上科目，避免混淆
    tmp = tmp.add_prefix(subject + '_')
    # 合併到總表
    if result.empty:
        result = tmp
    else:
        result = pd.concat([result, tmp], axis=1)

# 重設索引方便檢視
result = result.reset_index()

# 只顯示及格率（百分比）
for subject in ['使用者在國文測驗的分數', '使用者在數學測驗的分數', '使用者在英文測驗的分數']:
    result[subject+'_及格率'] = result[subject+'_及格率'] * 100

print(result)

# 欄位名稱對照表
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

# 重新命名欄位
result = result.rename(columns=rename_dict)

# 輸出 Excel 檔案
result.to_excel('學校分析結果.xlsx', index=False)


# 輸出 Excel 檔案
excel_path = '學校分析結果.xlsx'
result.to_excel(excel_path, index=False)

# 讀取剛剛寫出的 Excel 檔，調整欄寬
wb = load_workbook(excel_path)
ws = wb.active

for column_cells in ws.columns:
    # 取得這個欄所有字串的最大長度
    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
    # 適當加一點空間（例如+2）
    ws.column_dimensions[column_cells[0].column_letter].width = length*2 
wb.save(excel_path)