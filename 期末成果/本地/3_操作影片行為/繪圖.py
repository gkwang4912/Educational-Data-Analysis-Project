import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import os

# 來源資料夾（存放所有影片的事件轉移機率矩陣）
input_dir = "轉移機率矩陣_各影片"
# 圖片輸出資料夾
output_dir = "事件轉移圖_各影片"
os.makedirs(output_dir, exist_ok=True)

# 讀取資料夾中的每一個檔案
for filename in os.listdir(input_dir):
    if filename.endswith(".xlsx"):
        filepath = os.path.join(input_dir, filename)

        # 讀取轉移機率矩陣
        df = pd.read_excel(filepath, index_col=0)

        # 建立有向圖
        G = nx.DiGraph()
        for from_state in df.index:
            for to_state in df.columns:
                prob = df.at[from_state, to_state]
                if prob > 0:
                    G.add_edge(from_state, to_state, weight=round(prob, 3))

        # 繪圖
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

        # 設定標題與儲存圖片
        title = filename.replace(".xlsx", "")
        plt.title(f"{title} - 事件轉移機率圖", fontsize=16)
        plt.tight_layout()

        # 儲存圖片
        output_path = os.path.join(output_dir, f"{title}.png")
        plt.savefig(output_path)
        plt.close()

print(f"✅ 所有事件轉移圖已儲存在資料夾：{output_dir}")
