# Excel Worksheet Grader

A simple tool for grading Excel worksheets that compares Y/N values in row 1.

## Files

- `autograder.py` - Core grading functionality
- `uploader.py` - Upload and grade individual files
- `batch.py` - Grade multiple files at once
- `solution.xlsx` - Reference solution file
- `Dockerfile` - For Coursera integration

## Usage

### Individual Grading

Grade a single Excel file:

```bash
python uploader.py student_file.xlsx
```

This will:
- Copy the file to the `uploads` folder
- Grade it against `solution.xlsx`
- Display the score and feedback
- Save feedback to a text file

### Batch Grading

Grade multiple Excel files:

```bash
# Grade all Excel files in a directory
python batch.py submissions/

# Grade specific files
python batch.py file1.xlsx file2.xlsx

# Grade all Excel files in current directory
python batch.py
```

This will:
- Grade all the specified Excel files
- Save feedback files in the `results` folder
- Generate summary reports in CSV and Excel formats
- Display statistics and a summary table

### Feedback Format

The grader provides minimal feedback:

```
Your score: 85.33%
You correctly matched 64 out of 75 cells.
```

## Coursera Integration

To use this with Coursera:

1. Update the `PART_ID` in `autograder.py`
2. Build the Docker image:
   ```
   docker build -t excel-grader .
   ```
3. Test locally with Coursera autograder tool
4. Upload to Coursera

## Requirements

- Python 3.6 or higher
- openpyxl package (`pip install openpyxl`)
- pandas package for batch grading (`pip install pandas`)