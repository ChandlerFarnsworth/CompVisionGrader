#!/usr/bin/python3
"""
Upload tool for Excel files that automatically grades them against a solution

Usage:
  python upload_and_grade.py [filename]
  
If no filename is provided, the script will prompt for one.
"""

import os
import sys
import shutil
import datetime
from pathlib import Path

# Import the grading function from the autograder
from excel_autograder import grade_excel_worksheet, generate_feedback

# Constants
UPLOAD_FOLDER = "uploads"
SOLUTION_FILE = "solution.xlsx"

def ensure_upload_folder():
    """Create the uploads folder if it doesn't exist"""
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    print(f"✓ Upload folder ready: {UPLOAD_FOLDER}")

def generate_unique_filename(original_filename):
    """Generate a unique filename with timestamp to avoid overwrites"""
    # Get current timestamp
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Extract base name and extension
    base_name = Path(original_filename).stem
    extension = Path(original_filename).suffix
    
    # Create new filename with timestamp
    return f"{base_name}_{timestamp}{extension}"

def copy_file_to_uploads(source_path):
    """Copy the file to the uploads folder with a unique name"""
    if not os.path.exists(source_path):
        print(f"✗ Error: File not found: {source_path}")
        return None
    
    # Verify it's an Excel file
    extension = Path(source_path).suffix.lower()
    if extension not in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
        print(f"✗ Error: File must be an Excel file (.xlsx, .xlsm, etc.): {source_path}")
        return None
    
    # Check if the file is already in the uploads folder
    if os.path.dirname(os.path.abspath(source_path)) == os.path.abspath(UPLOAD_FOLDER):
        print(f"✓ File is already in uploads folder: {source_path}")
        return source_path
    
    # Generate a unique filename
    dest_filename = generate_unique_filename(os.path.basename(source_path))
    dest_path = os.path.join(UPLOAD_FOLDER, dest_filename)
    
    # Copy the file
    try:
        shutil.copyfile(source_path, dest_path)
        print(f"✓ File uploaded to: {dest_path}")
        return dest_path
    except Exception as e:
        print(f"✗ Error copying file: {e}")
        return None

def grade_uploaded_file(file_path):
    """Grade the uploaded file against the solution"""
    if not os.path.exists(SOLUTION_FILE):
        print(f"✗ Error: Solution file not found: {SOLUTION_FILE}")
        print("Please run extract_solution.py to create it first.")
        return
    
    print(f"\n===== GRADING: {os.path.basename(file_path)} =====")
    print(f"Comparing against solution: {SOLUTION_FILE}")
    print(f"Full file path: {os.path.abspath(file_path)}")
    
    # Verify the file exists
    if not os.path.exists(file_path):
        print(f"✗ Error: File does not exist at path: {file_path}")
        return
    
    # Print file info
    print(f"File size: {os.path.getsize(file_path)} bytes")
    print(f"File exists check: {os.path.exists(file_path)}")
    
    # Grade the submission
    try:
        result = grade_excel_worksheet(file_path, SOLUTION_FILE)
        
        # Display results
        if 'score' in result:
            percentage = result['score'] * 100
            print(f"\nSCORE: {percentage:.2f}%")
            
            # Create a feedback file
            feedback_path = file_path.replace('.xlsx', '_feedback.txt').replace('.xlsm', '_feedback.txt')
            with open(feedback_path, 'w') as f:
                f.write(result['feedback'])
            
            print(f"Detailed feedback saved to: {feedback_path}")
            
            # Print the feedback to the console
            print("\n" + "=" * 50)
            print(result['feedback'])
            print("=" * 50)
        else:
            print(f"Error grading file: {result.get('error', 'Unknown error')}")
    except Exception as e:
        print(f"Exception while grading: {str(e)}")

def main():
    """Main function to handle file upload and grading"""
    print("\n===== EXCEL WORKSHEET UPLOADER & GRADER =====")
    
    # Ensure upload folder exists
    ensure_upload_folder()
    
    # Get the file path from command-line args or prompt
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("\nEnter the path to the Excel file to upload: ")
    
    # Check if the file is already in the uploads folder
    if os.path.dirname(os.path.abspath(file_path)) == os.path.abspath(UPLOAD_FOLDER):
        print(f"File is already in uploads folder, using it directly")
        uploaded_file = file_path
    else:
        # Copy the file to uploads
        uploaded_file = copy_file_to_uploads(file_path)
    
    if uploaded_file:
        # Grade the uploaded file
        grade_uploaded_file(uploaded_file)
        
        print("\n✓ Process completed successfully!")
    else:
        print("\n✗ Upload failed. Please check the file and try again.")

if __name__ == "__main__":
    main()