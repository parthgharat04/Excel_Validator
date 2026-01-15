# 📊 Excel Stats Analyzer

A powerful desktop GUI application that analyzes Excel files and calculates comprehensive statistics for each column/header across multiple sheets.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## ✨ Features

### 📁 Multi-Sheet Support
- Load Excel files with multiple sheets (.xlsx, .xls, .xlsm)
- Select individual sheets, multiple sheets, or use **Select All** for batch processing
- Visual sheet selection with checkboxes

### 📈 Column Statistics
For each header/column in your Excel sheets, the application calculates:

| Statistic | Description |
|-----------|-------------|
| **Total Number of Transactions** | Total row count (includes blank values) |
| **Count of Availability** | Count of non-blank values only |
| **% Availability** | Percentage of non-blank values `(Count / Total × 100)` |
| **No of Unique Values** | Count of distinct non-blank values |

### 📋 Flexible Output Formats
Choose between two output modes:

1. **Separate Sheets Mode**
   - Each analyzed sheet gets its own tab in the output file
   - Clean, organized output per sheet

2. **Consolidated Sheet Mode**
   - All statistics combined into a single "Consolidated Stats" tab
   - Includes a "Sheet Name" column to identify the source sheet
   - Perfect for cross-sheet analysis and reporting

### 🛡️ Smart File Handling
- **Automatic file naming**: Output saved as `<input_filename>_stats.xlsx`
- **Conflict resolution**: If file exists, automatically appends incrementing numbers (`_stats_1.xlsx`, `_stats_2.xlsx`, etc.)
- **Large file support**: Efficient processing using pandas with background threading
- **Progress tracking**: Real-time progress bar and status updates

### 🎨 Modern UI
- Clean, dark-themed interface
- Responsive design with scrollable sheet list
- Clear step-by-step workflow

---

## 🚀 Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/parthgharat04/Excel_Validator.git
   cd Excel_Validator
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python3 -m venv venv
   
   # Activate on macOS/Linux
   source venv/bin/activate
   
   # Activate on Windows
   venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

---

## 📖 Usage

### Running the Application

```bash
# Make sure virtual environment is activated
./venv/bin/python excel_stats_analyzer.py

# Or on Windows
venv\Scripts\python excel_stats_analyzer.py
```

### Step-by-Step Guide

1. **Step 1: Select Excel File**
   - Click **Browse...** to open a file dialog
   - Select your Excel file (.xlsx, .xls, or .xlsm)

2. **Step 2: Select Sheets to Analyze**
   - Check individual sheets you want to analyze
   - Or click **Select All** to analyze all sheets

3. **Step 3: Choose Output Format**
   - **Separate Sheets**: Each sheet's stats in its own tab
   - **Consolidated Sheet**: All stats in a single tab with sheet identification

4. **Step 4: Generate Report**
   - Click **🔍 Analyze & Generate Report**
   - Monitor progress via the progress bar
   - Output file is saved in the same directory as input

---

## 📊 Output Examples

### Separate Sheets Mode

Each sheet creates its own tab. For a sheet named "Sales":

| Header Name | Total Number of Transactions | Count of Availability | % Availability | No of Unique Values |
|-------------|------------------------------|----------------------|----------------|---------------------|
| Product     | 1000                         | 998                  | 99.8           | 45                  |
| Price       | 1000                         | 1000                 | 100.0          | 120                 |
| Region      | 1000                         | 950                  | 95.0           | 8                   |

### Consolidated Sheet Mode

All sheets combined with source identification:

| Sheet Name | Header Name | Total Number of Transactions | Count of Availability | % Availability | No of Unique Values |
|------------|-------------|------------------------------|----------------------|----------------|---------------------|
| Sales      | Product     | 1000                         | 998                  | 99.8           | 45                  |
| Sales      | Price       | 1000                         | 1000                 | 100.0          | 120                 |
| Inventory  | SKU         | 500                          | 500                  | 100.0          | 500                 |
| Inventory  | Quantity    | 500                          | 485                  | 97.0           | 50                  |

---

## 🛠️ Troubleshooting

| Issue | Solution |
|-------|----------|
| **File won't load** | Ensure the file is a valid Excel format and not corrupted |
| **Permission denied** | Close the input file if it's open in Excel or another application |
| **Module not found** | Make sure virtual environment is activated and dependencies are installed |
| **Application not visible** | Check your taskbar/dock - the window may be minimized |

---

## 📦 Dependencies

- **pandas** >= 2.0.0 - Data processing and analysis
- **openpyxl** >= 3.1.0 - Excel file reading/writing (.xlsx)
- **xlrd** >= 2.0.0 - Legacy Excel file support (.xls)

---

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest new features
- Submit pull requests

---

## 📄 License

This project is open source and available under the MIT License.
