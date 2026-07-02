FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY tools /app/tools
COPY webapp /app/webapp

ENV SCM_ROOT=/app
ENV SCM_DATA_DIR=/data
ENV PORT=8080

EXPOSE 8080

CMD ["python", "/app/webapp/scm_web_dashboard.py"]
