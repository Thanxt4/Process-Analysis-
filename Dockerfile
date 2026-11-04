FROM python:3.12-slim
ENV PYTHONUNBUFFERED=1 DATA_DIR=/data TZ=Asia/Bangkok
WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app
WORKDIR /srv/app
VOLUME ["/data"]
EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
