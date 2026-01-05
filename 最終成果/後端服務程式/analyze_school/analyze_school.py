from flask import Flask, request, jsonify
import pandas as pd
from sqlalchemy import create_engine, text
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def analyze_school():
    try:
        # 連線設定
        conn_str = "postgresql+psycopg2://postgres:eric1611@35.229.163.253:5432/postgres"
        engine = create_engine(conn_str)

        # 讀取資料
        df = pd.read_sql('SELECT * FROM document."學期成績"', con=engine)

        # 數據轉換
        subjects = ['使用者在國文測驗的分數', '使用者在數學測驗的分數', '使用者在英文測驗的分數']
        for subject in subjects:
            df[subject] = pd.to_numeric(df[subject], errors='coerce')

        pass_score = 60
        grouped = df.groupby('學校代碼')
        result = pd.DataFrame()

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

        for subject in subjects:
            result[subject + '_及格率'] = result[subject + '_及格率'] * 100

        # 欄位重新命名
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

        # 調整欄寬
        wb = load_workbook(excel_path)
        ws = wb.active
        for column_cells in ws.columns:
            length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length * 2
        wb.save(excel_path)

        # 寫入資料庫
        table_name = os.path.splitext(os.path.basename(excel_path))[0]
        with engine.connect() as conn:
            conn.execute(text("CREATE SCHEMA IF NOT EXISTS result;"))
        result.to_sql(name=table_name, con=engine, schema='result', if_exists='replace', index=False)

        return jsonify({"status": "ok", "message": "學校分析完成，已寫入 Excel 與資料表"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
