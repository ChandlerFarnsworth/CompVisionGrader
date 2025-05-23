# Use Python 3.9 base image
FROM python:3.9-slim

# Set working directory
WORKDIR /grader

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements file
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy grader files
COPY . .

# Create necessary directories
RUN mkdir -p /shared/submission
RUN mkdir -p /grader/solutions
RUN mkdir -p /grader/output

# Set environment variables
ENV PYTHONPATH=/grader
ENV SUBMISSION_DIR=/shared/submission

# Make the grader script executable
RUN chmod +x grader.py

# Default command
CMD ["python3", "grader.py"]