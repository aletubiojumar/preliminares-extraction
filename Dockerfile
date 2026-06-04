FROM python:3.12-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1

# Módulos compartidos (browser, config, epac.pages.*) buscados en parent.parent = /
COPY preliminares-upload/ /preliminares-upload/

COPY preliminares-extraction/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && \
    playwright install chromium --with-deps

COPY preliminares-extraction/ .

CMD ["python", "bot_telegram.py"]
