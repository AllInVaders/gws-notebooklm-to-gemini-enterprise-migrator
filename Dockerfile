FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
ENV NBLM_WORKSPACE_DIR=/tmp/nblm_migration_workspace
EXPOSE 8765

CMD ["python3", "gws_to_ge_notebooklm_migrator.py", "--serve", "--port", "8765"]
