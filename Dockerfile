# Use Ubuntu 22.04 as base (more compatible with Coursera)
FROM ubuntu:22.04

# Set non-interactive mode for apt
ENV DEBIAN_FRONTEND=noninteractive

# Install Python 3.10, pip, and other essentials
RUN apt-get update && apt-get install -y \
    python3.10 \
    python3-pip \
    python3.10-venv \
    python3-distutils \
    curl \
    ca-certificates \
 && apt-get clean && rm -rf /var/lib/apt/lists/*

# Symlink python3.10 to python
RUN ln -s /usr/bin/python3.10 /usr/bin/python

# Install required Python packages
RUN python3.10 -m pip install --no-cache-dir pandas numpy openpyxl chardet

# Create grader directory
RUN mkdir /grader

# Copy files individually (more reliable for Coursera)
COPY autograder.py /grader/autograder.py
COPY grader.py /grader/grader.py
COPY solution.xlsx /grader/solution.xlsx

# Set proper permissions
RUN chmod a+rwx -R /grader/

# Make the autograder executable
RUN chmod +x /grader/autograder.py

# Set the entrypoint
ENTRYPOINT ["python3", "/grader/autograder.py"]