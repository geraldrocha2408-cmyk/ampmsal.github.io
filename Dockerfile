FROM python:3.11-slim

WORKDIR /app

COPY backend/requirements.txt /app/backend/requirements.txt
RUN pip install --no-cache-dir -r /app/backend/requirements.txt

COPY backend /app/backend
COPY data_api/out /app/data_api/out

ENV PYTHONUNBUFFERED=1

CMD python /app/backend/app/main.py
