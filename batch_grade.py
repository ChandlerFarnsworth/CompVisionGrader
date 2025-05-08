#!/usr/bin/python3
"""
Batch grading tool for Excel files

This script processes all Excel files in a specified folder (or specified files)
and grades them against the solution file.

Usage:
  python batch_grade.py [folder_path_or_file1 file2 ...]
  
If no arguments are provided, it processes all Excel files in the current directory.
"""

import os
import sys
import glob
import time
import pandas as pd
from pathlib import Path

# Import the grading function from the autograder
from excel_autograder import grade_excel_worksheet, generate_feedback

# Constants
SOLUTION_FILE = "solution.xlsx"
RESULTS_FOLDER = "results"

def ensure_results_folder():
    """Create the results folder if it doesn't exist"""
    os.makedirs(RESULTS_FOLDER, exist_ok=True)
    print(f"✓ Results folder ready: {RESULTS_FOLDER}")

def get_files_to_grade(args):
    """Get list of Excel files to grade based on arguments"""
    files_to_grade = []
    
    if not args:
        # No arguments - process all Excel files in current directory
        print("No files specified. Processing all Excel files in current directory.")
        files_to_grade = glob.glob("*.xlsx") + glob.glob("*.xlsm")
    else:
        for arg in args:
            if os.path.isdir(arg):
                # Argument is a directory - process all Excel files in it
                print(f"Processing directory: {arg}")
                dir_files = glob.glob(os.path.join(arg, "*.xlsx")) + glob.glob(os.path.join(arg, "*.xlsm"))
                files_to_grade.extend(dir_files)
            elif os.path.isfile(arg) and arg.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
                # Argument is an Excel file
                files_to_grade.append(arg)
            else:
                print(f"Warning: Skipping '{arg}' - not an Excel file or directory")
    
    return files_to_grade

def grade_file(file_path):
    """Grade an Excel file against the solution and return results"""
    print(f"Grading: {file_path}")
    
    try:
        result = grade_excel_worksheet(file_path, SOLUTION_FILE)
        
        if 'score' in result:
            return {
                'filename': os.path.basename(file_path),
                'path': file_path,
                'score': result['score'],
                'percentage': result['score'] * 100,
                'matches': result.get('matches', 0),
                'total': result.get('total_cells', 0),
                'feedback': result['feedback'],
                'status': 'Success'
            }
        else:
            return {
                'filename': os.path.basename(file_path),
                'path': file_path,
                'score': 0,
                'percentage': 0,
                'matches': 0,
                'total': 0,
                'feedback': result.get('feedback', 'Unknown error'),
                'status': 'Error'
            }
    except Exception as e:
        return {
            'filename': os.path.basename(file_path),
            'path': file_path,
            'score': 0,
            'percentage': 0,
            'matches': 0,
            'total': 0,
            'feedback': f"Error processing file: {str(e)}",
            'status': 'Error'
        }

def save_feedback_files(results):
    """Save individual feedback files for each submission"""
    for result in results:
        feedback_filename = Path(result['path']).stem + "_feedback.txt"
        feedback_path = os.path.join(RESULTS_FOLDER, feedback_filename)
        
        with open(feedback_path, 'w') as f:
            f.write(result['feedback'])

def generate_summary_report(results):
    """Generate a summary report of all graded files"""
    # Create a DataFrame from the results
    df = pd.DataFrame(results)
    
    # Create a summary report
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(RESULTS_FOLDER, f"grading_summary_{timestamp}.csv")
    
    # Select and reorder columns for the report
    report_columns = ['filename', 'percentage', 'matches', 'total', 'status']
    report_df = df[report_columns].copy()
    
    # Format percentage
    report_df['percentage'] = report_df['percentage'].apply(lambda x: f"{x:.2f}%")
    
    # Save the report
    report_df.to_csv(report_path, index=False)
    print(f"✓ Summary report saved to: {report_path}")
    
    # Also save an Excel version
    excel_path = os.path.join(RESULTS_FOLDER, f"grading_summary_{timestamp}.xlsx")
    report_df.to_excel(excel_path, index=False)
    print(f"✓ Excel summary saved to: {excel_path}")
    
    return report_df

def print_summary_table(results_df):
    """Print a formatted summary table to the console"""
    if len(results_df) == 0:
        print("No files were graded.")
        return
    
    print("\n===== GRADING SUMMARY =====")
    print(f"Total files processed: {len(results_df)}")
    
    # Calculate statistics
    score_values = [result['percentage'] for result in results 
                   if isinstance(result['percentage'], (int, float))]
    
    if score_values:
        avg_score = sum(score_values) / len(score_values)
        max_score = max(score_values)
        min_score = min(score_values)
        
        print(f"Average score: {avg_score:.2f}%")
        print(f"Highest score: {max_score:.2f}%")
        print(f"Lowest score: {min_score:.2f}%")
    
    # Print the table
    print("\n" + "=" * 80)
    print(f"{'Filename':<30} {'Score':<10} {'Matches':<15} {'Status':<10}")
    print("-" * 80)
    
    for _, row in results_df.iterrows():
        print(f"{row['filename']:<30} {row['percentage']:<10} {row['matches']}/{row['total']:<15} {row['status']:<10}")
    
    print("=" * 80)

def main():
    """Main function to handle batch grading"""
    print("\n===== EXCEL WORKSHEET BATCH GRADER =====")
    
    # Ensure solution file exists
    if not os.path.exists(SOLUTION_FILE):
        print(f"✗ Error: Solution file not found: {SOLUTION_FILE}")
        print("Please run extract_solution.py to create it first.")
        return 1
    
    # Ensure results folder exists
    ensure_results_folder()
    
    # Get files to grade
    files_to_grade = get_files_to_grade(sys.argv[1:])
    
    if not files_to_grade:
        print("No Excel files found to grade.")
        return 1
    
    print(f"\nFound {len(files_to_grade)} files to grade.")
    
    # Grade each file
    results = []
    for file_path in files_to_grade:
        result = grade_file(file_path)
        results.append(result)
    
    # Save individual feedback files
    save_feedback_files(results)
    
    # Generate and display summary report
    results_df = generate_summary_report(results)
    print_summary_table(results_df)
    
    print("\n✓ Batch grading completed successfully!")
    return 0

if __name__ == "__main__":
    sys.exit(main())