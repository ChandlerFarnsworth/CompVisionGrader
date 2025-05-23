FROM python:3.9-slim

# Debug: List all files in the current directory
RUN echo "=== DEBUGGING: Files in build context ===" && ls -la

WORKDIR /grader

# Debug: List files after setting workdir
RUN echo "=== DEBUGGING: Files in /grader ===" && ls -la

RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /shared/submission
RUN mkdir -p /grader/solutions
RUN mkdir -p /grader/output

ENV PYTHONPATH=/grader
ENV SUBMISSION_DIR=/shared/submission

RUN chmod +x grader.py

CMD ["python3", "autograder.py"]