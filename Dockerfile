FROM python:3.10-slim

# Instala o Tesseract e o Poppler no sistema
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . /app

# Instala as bibliotecas do Python
RUN pip install --no-cache-dir -r requirements.txt

# Roda o Streamlit
EXPOSE 8501
CMD ["streamlit", "run", "Menu_principal.py", "--server.port=8501", "--server.address=0.0.0.0"]
