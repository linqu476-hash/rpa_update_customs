# For hosting on a server INSIDE the company network (it must reach ax.d365ffo.onprem.libgroup.com)
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && playwright install --with-deps chromium
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
