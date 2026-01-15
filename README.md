# Excel Stats Analyzer

A desktop GUI application that analyzes Excel files and calculates statistics for each column/header across multiple sheets.

## Features

- **Multi-sheet support**: Select individual sheets, multiple sheets, or all sheets at once
- **Column statistics**: Calculates the following for each header:
  - **Total Number of Transactions**: Total row count (includes blank values)
  - **Count of Availability**: Count of non-blank values
  - **% Availability**: Percentage of non-blank values
  - **No of Unique Values**: Count of unique non-blank values
- **Large file handling**: Uses pandas for efficient processing
- **Clean output**: Generates a new Excel file with statistics for each sheet on a separate tab

## Installation

### Prerequisites
- Python 3.8 or higher

### Setup

1. Navigate to the project directory:
   ```bash
   cd /Users/parth.gharat/Development/Code/Excel_Validator
   ```

2. Create a virtual environment (recommended):
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On macOS/Linux
   # OR
   venv\Scripts\activate  # On Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python excel_stats_analyzer.py
   ```

2. Click **Browse...** to select your Excel file (.xlsx, .xls, or .xlsm)

3. Select the sheets you want to analyze:
   - Check individual sheets
   - Use **Select All** to select all sheets at once

4. Click **Analyze & Generate Report**

5. The output file will be saved in the same directory as the input file with `_stats` appended to the filename
   - Example: `mydata.xlsx` → `mydata_stats.xlsx`

## Output Format

For each selected sheet, the output Excel file will contain a corresponding sheet with statistics:

| Header Name | Total Number of Transactions | Count of Availability | % Availability | No of Unique Values |
|-------------|------------------------------|----------------------|----------------|---------------------|
| Column1     | 100                          | 95                   | 95.0           | 45                  |
| Column2     | 100                          | 100                  | 100.0          | 10                  |
| ...         | ...                          | ...                  | ...            | ...                 |

## Troubleshooting

- **File won't load**: Ensure the file is a valid Excel format (.xlsx, .xls, .xlsm) and is not corrupted
- **Permission denied**: Make sure the input file is not open in another application
- **Output file error**: Ensure you have write permissions in the input file's directory

