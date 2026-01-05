import pandas as pd

# 載入資料
df = pd.read_excel("作答狀況.xlsx", sheet_name="Sheet1")

# 將作答結果轉為 list
def parse_result(s):
    if isinstance(s, str):
        return [int(i) for i in s.split("@XX@") if i.strip().isdigit()]
    return []

df["小題結果list"] = df["練習題中每一小題的作答結果"].apply(parse_result)

# 統計基本資訊
summary = df.groupby("科目名稱").agg(
    作答次數=("使用者流水號", "count"),
    平均正確率=("練習題正確率", "mean"),
    平均作答時間_秒=("進行練習題的作答時間", "mean"),
)

# 每科目實際題數
question_counts = df.groupby("科目名稱")["小題結果list"].apply(lambda x: max(len(lst) for lst in x)).to_frame(name="實際題數")

# 每小題正確率統計
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

# 轉為 DataFrame
accuracy_df = pd.DataFrame.from_dict(subject_question_stats, orient="index")
accuracy_df.columns = [f"第{i+1}題正確率" for i in range(accuracy_df.shape[1])]
accuracy_df.index.name = "科目名稱"

# 合併所有資料（基本統計 + 題數 + 每題正確率）
final_df = pd.concat([summary, question_counts, accuracy_df], axis=1)

# 輸出 Excel
final_df.to_excel("練習題.xlsx")

print("✅ 已輸出：練習題.xlsx")



# 輸出 Excel
excel_path = "練習題.xlsx"
final_df.to_excel(excel_path)

# 加這段讓每個欄位寬度自動根據內容調整（需 openpyxl）
from openpyxl import load_workbook

wb = load_workbook(excel_path)
ws = wb.active

for column_cells in ws.columns:
    # 計算每個欄最大字串長度
    length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length*2

wb.save(excel_path)

print("✅ 已輸出：練習題.xlsx（欄位寬度已自動調整）")


