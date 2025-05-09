#!/usr/bin/python3
"""Batch grading tool for Excel files"""

import os
import sys
import glob
import time
import pandas as pd
from pathlib import Path

# Import the grading function from the local grader
from grader import grade_excel_worksheet

# Constants
RESULTS_FOLDER = "results"

def batch_grade(file_paths):
    """Grade multiple Excel files and generate reports"""
    # Create results folder
    os.makedirs(RESULTS_FOLDER, exist_ok=True)
    
    if not file_paths:
        print("No files to grade.")
        return False
    
    print(f"Found {len(file_paths)} files to grade.")
    
    # Grade each file
    results = []
    for path in file_paths:
        print(f"Grading: {path}")
        try:
            result = grade_excel_worksheet(path)
            
            # Store result
            if 'score' in result:
                results.append({
                    'filename': os.path.basename(path),
                    'path': path,
                    'score': result['score'],
                    'percentage': result['score'] * 100,
                    'matches': result.get('matches', 0),
                    'total': result.get('total_cells', 0),
                    'feedback': result['feedback'],
                    'status': 'Success'
                })
                
                # Save feedback
                feedback_path = os.path.join(RESULTS_FOLDER, 
                                           Path(path).stem + "_feedback.txt")
                with open(feedback_path, 'w') as f:
                    f.write(result['feedback'])
            else:
                results.append({
                    'filename': os.path.basename(path),
                    'path': path,
                    'score': 0,
                    'percentage': 0,
                    'matches': 0,
                    'total': 0,
                    'feedback': result.get('feedback', 'Unknown error'),
                    'status': 'Error'
                })
        except Exception as e:
            print(f"Error processing {path}: {e}")
            results.append({
                'filename': os.path.basename(path),
                'path': path,
                'score': 0,
                'percentage': 0,
                'matches': 0,
                'total': 0,
                'feedback': f"Error: {str(e)}",
                'status': 'Error'
            })
    
    # Create report if we have results
    if results:
        # Generate reports
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        report_name = f"grading_summary_{timestamp}"
        
        # Create DataFrame
        df = pd.DataFrame(results)
        columns = ['filename', 'percentage', 'matches', 'total', 'status']
        report_df = df[columns].copy()
        report_df['percentage'] = report_df['percentage'].apply(lambda x: f"{x:.2f}%")
        
        # Save reports
        report_df.to_csv(os.path.join(RESULTS_FOLDER, f"{report_name}.csv"), index=False)
        report_df.to_excel(os.path.join(RESULTS_FOLDER, f"{report_name}.xlsx"), index=False)
        
        # Print summary
        print("\n===== GRADING SUMMARY =====")
        print(f"Total files processed: {len(results)}")
        
        # Calculate statistics
        scores = [r['percentage'] for r in results if isinstance(r['percentage'], (int, float))]
        if scores:
            print(f"Average score: {sum(scores)/len(scores):.2f}%")
            print(f"Highest score: {max(scores):.2f}%")
            print(f"Lowest score: {min(scores):.2f}%")
        
        # Print table
        print("\n" + "=" * 80)
        print(f"{'Filename':<30} {'Score':<10} {'Matches':<15} {'Status':<10}")
        print("-" * 80)
        
        for _, row in report_df.iterrows():
            print(f"{row['filename']:<30} {row['percentage']:<10} {row['matches']}/{row['total']:<15} {row['status']:<10}")
        
        print("=" * 80)
        print(f"\nReports saved to: {RESULTS_FOLDER}/{report_name}.csv/xlsx")
        return True
    
    return False

def main():
    print("\n===== EXCEL WORKSHEET BATCH GRADER =====")
    
    # Get files to grade
    files_to_grade = []
    args = sys.argv[1:]
    
    if not args:
        # No arguments - process all Excel files in current directory
        print("Processing all Excel files in current directory.")
        files_to_grade = glob.glob("*.xlsx") + glob.glob("*.xlsm")
    else:
        for arg in args:
            if os.path.isdir(arg):
                # Process all Excel files in directory
                print(f"Processing directory: {arg}")
                files_to_grade.extend(glob.glob(os.path.join(arg, "*.xlsx")))
                files_to_grade.extend(glob.glob(os.path.join(arg, "*.xlsm")))
            elif os.path.isfile(arg) and arg.lower().endswith(('.xlsx', '.xlsm')):
                files_to_grade.append(arg)
            else:
                print(f"Skipping '{arg}' - not an Excel file or directory")
    
    # Run batch grading
    if batch_grade(files_to_grade):
        print("\nBatch grading completed successfully!")
    else:
        print("\nBatch grading failed or had no files to process.")

if __name__ == "__main__":
    main()