FROM python:3.12-slim

WORKDIR /app

# Install Python dependencies (no easyocr to keep image small)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Render sets PORT dynamically
ENV PORT=8000
EXPOSE ${PORT}

# Start server using Render's PORT
CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT} --timeout-keep-alive 300
