import pandas as pd
import os

# 讀取資料
file_path = "操作影片行為.xlsx"
df = pd.read_excel(file_path)

# 保留欄位與清理時間格式
df = df[["影片瀏覽流水號", "執行影片操作的時間戳記", "影片操作的行為名稱"]]
df = df[df["執行影片操作的時間戳記"] != "view_time"]
df["執行影片操作的時間戳記"] = pd.to_datetime(df["執行影片操作的時間戳記"], errors="coerce")
df = df.dropna(subset=["執行影片操作的時間戳記"])

# 創建輸出資料夾
output_dir = "轉移機率矩陣_各影片"
os.makedirs(output_dir, exist_ok=True)

# 根據每種影片瀏覽流水號進行分析
unique_ids = df["影片瀏覽流水號"].unique()

for vid in unique_ids:
    sub_df = df[df["影片瀏覽流水號"] == vid].sort_values(by="執行影片操作的時間戳記").copy()

    # 建立下一個行為欄位
    sub_df["下一個行為"] = sub_df["影片操作的行為名稱"].shift(-1)
    sub_df = sub_df.dropna(subset=["下一個行為"])

    # 統計轉移次數
    trans_counts = sub_df.groupby(["影片操作的行為名稱", "下一個行為"]).size().unstack(fill_value=0)

    # 轉換為機率
    trans_probs = trans_counts.div(trans_counts.sum(axis=1), axis=0)

    # 輸出每部影片的轉移機率矩陣
    filename = f"{output_dir}/影片_{vid}_轉移機率矩陣.xlsx"
    trans_probs.to_excel(filename)

print(f"✅ 分析完成，已為 {len(unique_ids)} 部影片輸出各自的轉移機率矩陣到資料夾：{output_dir}")
