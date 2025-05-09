# Base image
FROM ubuntu:22.04

# Set non-interactive mode for apt
ENV DEBIAN_FRONTEND=noninteractive

# Install Python and dependencies
RUN apt-get update && apt-get install -y \
    python3.10 \
    python3-pip \
    python3-distutils \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Install openpyxl
RUN pip install openpyxl

# Create grader directory
RUN mkdir /grader

# Copy autograder and solution
COPY autograder.py /grader/autograder.py
COPY solution.xlsx /grader/solution.xlsx

# Set permissions
RUN chmod a+x /grader/autograder.py

# Set entrypoint
ENTRYPOINT ["/grader/autograder.py"]