#!/usr/bin/python3
"""Simple uploader for grading Excel files"""

import os
import sys
import shutil
from datetime import datetime
from pathlib import Path

# Import grading function from local grader module
from grader import grade_excel_worksheet

# Constants
UPLOAD_FOLDER = "uploads"

def upload_and_grade(file_path):
    """Upload and grade an Excel file"""
    # Create upload folder if needed
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return False
    
    # Check if it's an Excel file
    if not file_path.lower().endswith(('.xlsx', '.xlsm')):
        print(f"Error: File must be an Excel file (.xlsx, .xlsm)")
        return False
        
    # Generate destination path
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = Path(file_path).name
    dest_path = os.path.join(UPLOAD_FOLDER, f"{Path(filename).stem}_{timestamp}{Path(filename).suffix}")
    
    # Skip copying if already in uploads folder
    if os.path.dirname(os.path.abspath(file_path)) == os.path.abspath(UPLOAD_FOLDER):
        dest_path = file_path
        print(f"File already in uploads folder")
    else:
        # Copy the file
        try:
            shutil.copyfile(file_path, dest_path)
            print(f"File uploaded to: {dest_path}")
        except Exception as e:
            print(f"Error copying file: {e}")
            return False
    
    # Grade the file
    print(f"\n===== GRADING: {Path(dest_path).name} =====")
    result = grade_excel_worksheet(dest_path)
    
    # Display results
    if 'score' in result:
        print("\n" + "=" * 50)
        print(result['feedback'])
        print("=" * 50)
        
        # Save feedback
        feedback_path = str(dest_path).rsplit('.', 1)[0] + "_feedback.txt"
        with open(feedback_path, 'w') as f:
            f.write(result['feedback'])
        print(f"Feedback saved to: {feedback_path}")
        return True
    else:
        print(f"Error grading file: {result.get('feedback', 'Unknown error')}")
        return False

def main():
    print("\n===== EXCEL WORKSHEET UPLOADER & GRADER =====")
    
    # Get file path
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Enter the path to the Excel file: ")
    
    # Process the file
    if upload_and_grade(file_path):
        print("\nProcess completed successfully!")
    else:
        print("\nUpload failed. Please check the file and try again.")

if __name__ == "__main__":
    main()