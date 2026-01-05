import pandas as pd
from sqlalchemy import create_engine
from openpyxl import load_workbook
import os
from sqlalchemy import text


# 資料庫連線字串（請保持正確）
conn_str = "postgresql+psycopg2://postgres:eric1611@35.229.163.253:5432/postgres"
engine = create_engine(conn_str)

# 從 PostgreSQL 的「作答狀況」資料表抓資料（注意需加雙引號）
df = pd.read_sql('SELECT * FROM document."作答狀況"', con=engine)


# 顯示欄位名稱（可幫助除錯）
print("📋 欄位名稱：", df.columns.tolist())

# 將作答結果欄位轉成 list（此欄名稱依你提供為被截斷版本）
def parse_result(s):
    if isinstance(s, str):
        return [int(i) for i in s.split("@XX@") if i.strip().isdigit()]
    return []

df["小題結果list"] = df["練習題中每一小題的作答結果"].apply(parse_result)

# 清洗數值欄位，避免計算平均時出錯
df["練習題正確率"] = pd.to_numeric(df["練習題正確率"], errors="coerce")
df["進行練習題的作答時間"] = pd.to_numeric(df["進行練習題的作答時間"], errors="coerce")

# 統計每科目基本資料
summary = df.groupby("科目名稱").agg(
    作答次數=("使用者流水號", "count"),
    平均正確率=("練習題正確率", "mean"),
    平均作答時間_秒=("進行練習題的作答時間", "mean"),
)

# 統計每科目實際題數
question_counts = df.groupby("科目名稱")["小題結果list"].apply(
    lambda x: max(len(lst) for lst in x)
).to_frame(name="實際題數")

# 每小題正確率計算
subject_question_stats = {}
for subject, group in df.groupby("科目名稱"):
    result_lists = group["小題結果list"].tolist()
    max_questions = max(len(lst) for lst in result_lists)

    correct_counts = [0] * max_questions
    total_counts = [0] * max_questions

    for result in result_lists:
        for i, val in enumerate(result):
            correct_counts[i] += val
            total_counts[i] += 1

    accuracy_per_question = [
        round(correct_counts[i] / total_counts[i] * 100, 2) if total_counts[i] > 0 else None
        for i in range(max_questions)
    ]

    subject_question_stats[subject] = accuracy_per_question

# 每題正確率轉為 DataFrame
accuracy_df = pd.DataFrame.from_dict(subject_question_stats, orient="index")
accuracy_df.columns = [f"第{i+1}題正確率" for i in range(accuracy_df.shape[1])]
accuracy_df.index.name = "科目名稱"

# 合併所有資料
final_df = pd.concat([summary, question_counts, accuracy_df], axis=1)

# 輸出 Excel
excel_path = "練習題.xlsx"
final_df.to_excel(excel_path)

# 自動調整欄寬（需 openpyxl）
wb = load_workbook(excel_path)
ws = wb.active
for column_cells in ws.columns:
    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length * 2
wb.save(excel_path)

print("✅ 成功分析並輸出 Excel：練習題.xlsx（欄寬自動調整）")


# 將 final_df 存入 PostgreSQL 的 result schema

# 取得表名（去除 .xlsx 副檔名）
table_name = os.path.splitext(os.path.basename(excel_path))[0]  # 會是 "練習題"

# 建立 result schema（若尚未存在）
with engine.connect() as conn:
    conn.execute(text("CREATE SCHEMA IF NOT EXISTS result;"))


# 寫入 result schema 中的資料表
final_df.to_sql(name=table_name, con=engine, schema='result', if_exists='replace', index=True)

print(f"📤 統計結果已上傳至 result.\"{table_name}\" 資料表")
