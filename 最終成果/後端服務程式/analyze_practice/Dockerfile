# 使用 Python 3.12 的官方映像檔
FROM python:3.12-slim

# 設定工作目錄
WORKDIR /app

# 複製檔案
COPY . .

# 安裝必要套件
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# 設定 Flask 的啟動點
ENV FLASK_APP=analyze_practice.py
ENV FLASK_RUN_PORT=8080

# 對外開放埠口
EXPOSE 8080

# 啟動 Flask，注意不要使用 debug 模式
CMD ["flask", "run", "--host=0.0.0.0", "--port=8080"]
