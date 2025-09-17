FROM python:3.11-slim

WORKDIR /app

# Install system deps for pandas/openpyxl if needed
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for caching
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . /app

# Create folders if they don't exist (runtime)
RUN mkdir -p uploads outputs

EXPOSE 5000

# Use gunicorn for production; keep Flask debug off
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]
