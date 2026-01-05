import pandas as pd
import psycopg2
import os

# PostgreSQL 連線
conn = psycopg2.connect(
    host="34.81.25.58",
    database="postgres",
    user="postgres",
    password="eric1611",
    port=5432
)
cur = conn.cursor()

# 建立兩個 schema（如果尚未存在）
for schema in ["document", "result"]:
    cur.execute(f"CREATE SCHEMA IF NOT EXISTS {schema};")
conn.commit()

# 定義檔案分配（你只要修改這兩個清單即可）
result_files = [
    "1_學校分析結果.xlsx",
    "2_練習題.xlsx",
    "3_影片操作轉移矩陣.xlsx"
]

document_files = [
    "作答狀況.xlsx"
]

# 合併成統一 dict（檔名 → schema）
file_schema_map = {file: "result" for file in result_files}
file_schema_map.update({file: "document" for file in document_files})

# 處理每個檔案
for file, schema in file_schema_map.items():
    print(f"處理中：{file} → schema: {schema}")
    df = pd.read_excel(file)

    # 資料表名稱：檔名（去副檔名）轉為小寫
    table_name = os.path.splitext(file)[0].lower()
    full_table_name = f'{schema}."{table_name}"'

    # 建立資料表
    columns = df.columns
    sql_fields = []
    for col in columns:
        colname = col.strip().replace(" ", "_").lower()
        sql_fields.append(f'"{colname}" TEXT')
    create_sql = f'CREATE TABLE IF NOT EXISTS {full_table_name} ({", ".join(sql_fields)});'

    cur.execute(create_sql)
    conn.commit()

    # 插入資料
    for _, row in df.iterrows():
        values = [str(v) if pd.notnull(v) else None for v in row.values]
        placeholders = ', '.join(['%s'] * len(values))
        insert_sql = f'INSERT INTO {full_table_name} VALUES ({placeholders})'
        cur.execute(insert_sql, values)
    conn.commit()

    print(f"✅ {file} 已匯入至 {full_table_name}")

cur.close()
conn.close()
print("🎉 所有檔案皆匯入完成！")
