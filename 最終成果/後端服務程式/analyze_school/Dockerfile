# 使用官方 Python 3.12 映像檔
FROM python:3.12-slim

# 設定工作目錄
WORKDIR /app

# 複製程式碼與需求檔案
COPY . /app

# 安裝相依套件
RUN pip install --upgrade pip \
 && pip install --no-cache-dir -r requirements.txt

# 暴露 Cloud Run 預設的 port
EXPOSE 8080

# 啟動應用（app.py 是你的 Python 檔名，如果叫 analyze_school.py 就改成那個）
CMD ["gunicorn", "-b", "0.0.0.0:8080", "analyze_school:app"]
