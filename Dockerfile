# Use official Python image
FROM python:3.12-slim

# Set working directory
WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY app.py ./
COPY templates ./templates

# Expose Flask port
EXPOSE 6969

# Run the Flask app
CMD ["python", "app.py"]
