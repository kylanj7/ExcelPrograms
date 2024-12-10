# Excel Tools Collection

A collection of Python scripts for manipulating and analyzing Excel files. These tools provide functionality for sorting, splitting, merging, creating pivot tables, and removing duplicates from Excel workbooks.

## Tools Overview

### 1. ExcelSort.py
Advanced Excel file sorting tool with features:
- Automatic header detection
- Support for multiple column types (text, dates, numbers)
- Case-insensitive column matching
- Detailed progress reporting
- Timestamped output files
- Comprehensive error handling
- Support for various product-related columns

### 2. ExcelSplitColumns.py
Split Excel files based on column values:
- Creates separate files for each unique value
- Maintains original data structure
- Automatic file naming based on split values

### 3. ExcelMerge.py
Combine multiple Excel files:
- Merges files while preserving structure
- Concatenates data vertically
- Maintains column headers
- Creates single consolidated output file

### 4. ExcelPivotTable.py
Create pivot tables from Excel data:
- Flexible index column selection
- Multiple value columns support
- Customizable aggregation functions
- Automated output generation

### 5. ExcelRMVduplicates.py
Remove duplicate entries:
- Multi-column duplicate detection
- Preserves first occurrence
- Creates cleaned output file
- Maintains data integrity

## Prerequisites

- Python 3.7+
- Required packages:
```bash
pip install pandas openpyxl
```

## Installation

1. Clone the repository:
```bash
git clone [your-repository-url]
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### ExcelSort.py
```python
python ExcelSort.py
# Default sorts "ITAD Recieving.xlsx" by "Item Name"
# Creates sorted file in SortedSheets folder
```

### ExcelSplitColumns.py
```python
python ExcelSplitColumns.py
# Splits "employees.xlsx" by "Department"
# Creates separate files for each department
```

### ExcelMerge.py
```python
python ExcelMerge.py
# Merges specified Excel files
# Creates "merged.xlsx"
```

### ExcelPivotTable.py
```python
python ExcelPivotTable.py
# Creates pivot table from "sales.xlsx"
# Outputs "pivot_sales.xlsx"
```

### ExcelRMVduplicates.py
```python
python ExcelRMVduplicates.py
# Removes duplicates from "duplicates.xlsx"
# Creates "cleaned_duplicates.xlsx"
```

## Features

### ExcelSort.py Features
- Valid column detection
- Automated folder creation
- Header row detection
- Date handling
- Detailed progress logging
- Error handling and reporting

### Common Features Across All Tools
- Non-destructive operations (creates new files)
- Error handling
- Progress reporting
- Data validation
- Memory efficient processing

## Output Locations

- ExcelSort.py: `C:\Users\KJohnson\Desktop\ExcelPrograms\SortedSheets\`
- Other tools: Same directory as input files with prefixes:
  - Split files: `[value]_`
  - Merged files: `merged_`
  - Pivot tables: `pivot_`
  - Cleaned files: `cleaned_`

## Error Handling

All scripts include error handling for:
- File access issues
- Invalid data formats
- Missing columns
- Memory constraints
- Permission errors

## Best Practices

1. Close Excel files before processing
2. Backup data before operations
3. Verify column names match exactly
4. Monitor available disk space
5. Check output files after processing

## Contributing

Feel free to submit issues and enhancement requests!

## License

[Specify your license here]

## Notes

- Large files may require additional memory
- Date formatting follows Excel standards
- Column names are case-sensitive (except in ExcelSort.py)
- Backup files before processing
