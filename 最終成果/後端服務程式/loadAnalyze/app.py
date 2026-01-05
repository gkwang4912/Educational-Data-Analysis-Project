from flask import Flask, jsonify
from flask_cors import CORS
from sqlalchemy import create_engine, text
import pandas as pd
import os

app = Flask(__name__)
CORS(app)  # ✅ 啟用 CORS，讓前端可跨網域請求

# PostgreSQL 連線字串（可改為使用環境變數）
# 原本的 IP 是 35.229.163.253，改為新的資料庫 IP（34.81.25.58）
conn_str = os.getenv(
    "DB_CONN",
    "postgresql+psycopg2://postgres:eric1611@34.81.25.58:5432/postgres"
)
engine = create_engine(conn_str)


@app.route("/")
def root():
    return "✅ API 正常運作中"

@app.route("/api/<schema>/<table>")
def get_table_data(schema, table):
    try:
        query = text(f'SELECT * FROM "{schema}"."{table}"')
        df = pd.read_sql(query, con=engine)
        df = df.fillna("")
        return jsonify(df.to_dict(orient="records"))
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
