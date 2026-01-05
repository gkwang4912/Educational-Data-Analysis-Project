import pandas as pd
from sqlalchemy import create_engine, text
import os

# PostgreSQL 連線設定
conn_str = "postgresql+psycopg2://postgres:eric1611@35.229.163.253:5432/postgres"
engine = create_engine(conn_str)

# 從 document schema 抓資料表 "操作影片行為"
df = pd.read_sql('SELECT * FROM document."操作影片行為"', con=engine)

# 清洗欄位與時間
df = df[["影片瀏覽流水號", "執行影片操作的時間戳記", "影片操作的行為名稱"]]
df = df[df["執行影片操作的時間戳記"] != "view_time"]
df["執行影片操作的時間戳記"] = pd.to_datetime(df["執行影片操作的時間戳記"], errors="coerce")
df = df.dropna(subset=["執行影片操作的時間戳記"])

# 分析每部影片的行為轉移
unique_ids = df["影片瀏覽流水號"].unique()
all_matrices = []

for vid in unique_ids:
    sub_df = df[df["影片瀏覽流水號"] == vid].sort_values(by="執行影片操作的時間戳記").copy()

    sub_df["下一個行為"] = sub_df["影片操作的行為名稱"].shift(-1)
    sub_df = sub_df.dropna(subset=["下一個行為"])

    # 統計轉移次數與機率
    trans_counts = sub_df.groupby(["影片操作的行為名稱", "下一個行為"]).size().unstack(fill_value=0)
    trans_probs = trans_counts.div(trans_counts.sum(axis=1), axis=0)

    # 轉為長格式，加上影片 ID
    matrix_long = trans_probs.reset_index().melt(
        id_vars="影片操作的行為名稱",
        var_name="下一個行為",
        value_name="轉移機率"
    )
    matrix_long["影片瀏覽流水號"] = vid
    all_matrices.append(matrix_long)

# 合併所有影片的轉移資料
final_df = pd.concat(all_matrices, ignore_index=True)
final_df["轉移機率"] = final_df["轉移機率"].fillna(0)

# 匯出 Excel
excel_path = "影片操作轉移矩陣.xlsx"
final_df.to_excel(excel_path, index=False)
print(f"✅ 已完成分析並輸出 Excel：{excel_path}")

# 上傳到 PostgreSQL 的 result schema
with engine.connect() as conn:
    conn.execute(text("CREATE SCHEMA IF NOT EXISTS result;"))

final_df.to_sql(name="影片操作轉移矩陣", con=engine, schema="result", if_exists="replace", index=False)
print('📤 已將轉移機率矩陣上傳至資料表：result."影片操作轉移矩陣"')
