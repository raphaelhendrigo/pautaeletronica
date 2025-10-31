# Dockerfile
FROM mcr.microsoft.com/playwright/python:v1.47.0-jammy

# Ajustes básicos
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    TZ=America/Sao_Paulo

WORKDIR /app

# Copia o projeto
COPY . /app

# Dependências Python (se tiver requirements.txt, descomente a linha correspondente)
# RUN pip install -r requirements.txt
RUN pip install --upgrade pip && \
    pip install \
      python-dotenv pandas openpyxl python-docx lxml numpy flask gunicorn \
      playwright

# Instala o browser Chromium
RUN python -m playwright install chromium

# Porta do Cloud Run
EXPOSE 8080

# Server HTTP para acionar o pipeline
CMD ["gunicorn", "-b", "0.0.0.0:8080", "server:app", "-w", "1", "-k", "gthread"]
