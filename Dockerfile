# Fetch ubuntu 22.04 LTS docker image
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

# Symlink python3.10 to python (if desired)
RUN ln -s /usr/bin/python3.10 /usr/bin/python

# Install openpyxl for Excel file processing
RUN python3.10 -m pip install --no-cache-dir openpyxl

# Create grader directory
RUN mkdir /grader

# Copy autograder files
COPY excel_autograder.py /grader/excel_autograder.py
COPY solution.xlsx /grader/solution.xlsx

# Set permissions
RUN chmod a+rwx -R /grader/

# Set the entrypoint to the autograder script
ENTRYPOINT [ "grader/excel_autograder.py" ]