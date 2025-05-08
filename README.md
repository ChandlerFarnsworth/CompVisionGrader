# Excel Worksheet Uploader & Grader - Usage Guide

This guide explains how to use the Excel uploader and grader tools included in this package. These tools allow you to upload Excel files, store them in a designated folder, and automatically grade them against a solution file.

## Available Tools

This package includes three different ways to upload and grade Excel files:

1. **Command-line Uploader** (`upload_and_grade.py`) - Simple command-line tool for uploading and grading individual files
2. **Batch Grader** (`batch_grade.py`) - Process multiple Excel files at once
3. **GUI Uploader** (`gui_upload_grade.py`) - Graphical interface for file selection and grading

## Prerequisites

Before using any of these tools, you need to:

1. Make sure you have Python 3.7+ installed
2. Install required packages:
   ```bash
   pip install openpyxl pandas
   ```
3. Create the solution file by running:
   ```bash
   python extract_solution.py Finalreal.xlsx
   ```

## 1. Command-line Uploader

The command-line uploader allows you to upload and grade a single Excel file.

### Usage

```bash
python upload_and_grade.py [file_path]
```

If you don't provide a file path, the script will prompt you to enter one.

### Example

```bash
python upload_and_grade.py student_submission.xlsx
```

This will:
1. Copy the file to the `uploads` folder with a timestamp in the filename
2. Grade it against `solution.xlsx`
3. Display the score and feedback
4. Save the feedback to a text file in the `uploads` folder

## 2. Batch Grader

The batch grader allows you to process multiple Excel files at once.

### Usage

```bash
python batch_grade.py [folder_or_files]
```

You can specify:
- A folder containing Excel files
- Multiple Excel files separated by spaces
- No arguments (to process all Excel files in the current directory)

### Examples

Process all Excel files in a specific folder:
```bash
python batch_grade.py student_submissions/
```

Process specific files:
```bash
python batch_grade.py file1.xlsx file2.xlsx file3.xlsx
```

Process all Excel files in the current directory:
```bash
python batch_grade.py
```

This will:
1. Grade each file against `solution.xlsx`
2. Create individual feedback files in the `results` folder
3. Generate a summary report in CSV and Excel formats
4. Display a summary table in the console

## 3. GUI Uploader

The GUI uploader provides a graphical interface for selecting, uploading, and grading Excel files.

### Usage

```bash
python gui_upload_grade.py
```

This will open a window where you can:
1. Click "Select Excel Files" to choose files from your computer
2. Click "Upload and Grade Files" to process them
3. View results in the text area
4. Save detailed feedback to text files
5. Open the uploads folder to access the processed files

## File Locations

- Uploaded files are stored in the `uploads` folder
- Batch processing results are stored in the `results` folder
- Feedback files are saved next to the uploaded files with `_feedback.txt` suffix

## Troubleshooting

If you encounter issues:

1. **Missing solution file**: Run `extract_solution.py` first
2. **Import errors**: Make sure you've installed required packages
3. **Permission errors**: Check that you have write permissions for the uploads and results folders
4. **Excel format issues**: Ensure the worksheet names match those expected by the autograder

## Customizing the Grader

If you need to adjust how the grading works:

1. The main grading logic is in `excel_autograder.py`
2. To change the expected worksheet names, modify the `STUDENT_SHEET_NAME` and `SOLUTION_SHEET_NAME` constants
3. To adjust the scoring or feedback, modify the `grade_excel_worksheet` and `generate_feedback` functions

---

For additional help or to report issues, please contact your instructor or teaching assistant.