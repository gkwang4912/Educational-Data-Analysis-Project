import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import os

# 讀取剛才的合併檔案
input_excel = "影片操作轉移矩陣.xlsx"

# 圖片輸出資料夾
output_dir = "事件轉移圖_各影片"
os.makedirs(output_dir, exist_ok=True)

# 讀取資料
df = pd.read_excel(input_excel)

# 確保資料正確
required_cols = ["影片瀏覽流水號", "影片操作的行為名稱", "下一個行為", "轉移機率"]
if not all(col in df.columns for col in required_cols):
    raise ValueError("❌ 欄位名稱不符，請確認 Excel 檔案格式")

# 依每部影片繪製事件轉移圖
for vid, sub_df in df.groupby("影片瀏覽流水號"):
    G = nx.DiGraph()
    for _, row in sub_df.iterrows():
        from_state = row["影片操作的行為名稱"]
        to_state = row["下一個行為"]
        weight = round(row["轉移機率"], 3)
        if weight > 0:
            G.add_edge(from_state, to_state, weight=weight)

    # 繪製圖形
    plt.figure(figsize=(12, 8))
    pos = nx.spring_layout(G, k=0.6, seed=42)

    nx.draw(G, pos,
            with_labels=True,
            node_size=2000,
            node_color="lightblue",
            edge_color="gray",
            font_size=10,
            font_weight="bold",
            arrowsize=20)

    edge_labels = nx.get_edge_attributes(G, 'weight')
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels)

    plt.title(f"影片 {vid} - 事件轉移圖", fontsize=16)
    plt.tight_layout()

    # 儲存圖檔
    filename = f"影片_{vid}_事件轉移圖.png"
    output_path = os.path.join(output_dir, filename)
    plt.savefig(output_path)
    plt.close()

print(f"✅ 所有影片事件轉移圖已完成並儲存在：{output_dir}")
