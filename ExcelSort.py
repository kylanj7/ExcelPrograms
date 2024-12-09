import pandas as pd
import os
from datetime import datetime

def get_valid_columns():
    """Returns list of all valid column names"""
    return [
        # Basic Product Information
        "Item Name",
        "Item",
        "Product Name",
        "Product Number",
        "PN",
        "Serial Number",
        "Batch Size",
        "Item Status",
        "Brand",
        "Manufacturer",
        "Make",
        "Model",
        "Type",
        "Type of Device",
        "Device Type",
        
        # Dates and Tracking
        "Date Aquired",
        "Date Posted",
        "Date Sold",
        "Test Date",
        "Return Date",
        
        # Sales and Inventory
        "Item Number",
        "Order Number",
        "QTY",
        "Price Sold",
        "Gross Sales",
        "Storage number",
        "Assigned Serial Number",
        
        # Returns and eBay
        "Returned(Y/N)",
        "Reason For Return",
        "Ebay Item ID",
        "eBay ID",
        "eBay Item #",
        "EID",
        
        # Technical Specifications
        "Chipset",
        "Power Cable",
        "Memory Type",
        "Memory Size",
        "Display Output",
        "Compatible Slot",
        "DC Resistance",
        "Rated Ω",
        
        # GPU Specific
        "GPU Model",
        "GPU Series",
        
        # Testing and Performance
        "Test Type",
        "Test Duration",
        "FPS/Score",
        "Degrees Celcius",
        "Pass Test?",
        "Audio Test Pass/Fail",
        "All A/V Input/ Output Tested",
        "Mixer/Volume Pots/Buttons Tested for Scratching",
        "Dante Network Pass",
        "Buttons Work",
        "Screen/Bulbs Pass Test",
        "Bulb Hours",
        
        # Quality Assessment
        "Cosmetic Grade",
        "Functional Grade",
        "Functionality Grade",
        "Cosmetic REC",
        "Functional REC",
        "QA Technician Name",
        "Qced by",
        "Quality Assurance",
        "Quality Assurance Testing",
        "Q A Check",
        "Tech",
        "Notes"
    ]

def create_sorted_folder():
    """Create SortedSheets folder if it doesn't exist"""
    base_path = r"C:\Users\KJohnson\Desktop\ExcelPrograms"
    sorted_folder = os.path.join(base_path, "SortedSheets")
    if not os.path.exists(sorted_folder):
        try:
            os.makedirs(sorted_folder)
            print(f"Created folder: {sorted_folder}")
        except Exception as e:
            print(f"Error creating folder: {e}")
    return base_path, sorted_folder

def detect_header_row(file_path):
    """Automatically detect which row contains the headers"""
    try:
        # Read first 10 rows to analyze
        preview_df = pd.read_excel(file_path, nrows=10)
        print("\nAnalyzing file structure...")
        
        for skip_rows in range(10):
            df = pd.read_excel(file_path, skiprows=skip_rows, nrows=5)
            unnamed_count = sum(1 for col in df.columns if 'Unnamed' in str(col))
            if unnamed_count < len(df.columns) * 0.5:
                print(f"Found headers at row {skip_rows + 1}")
                return skip_rows
        return 0
    except Exception as e:
        print(f"Error detecting headers: {e}")
        return 0

def sort_excel_file(input_filename, sort_by_column):
    """Sort Excel file by specified column"""
    try:
        # Setup paths
        base_path, sorted_folder = create_sorted_folder()
        input_file = os.path.join(base_path, input_filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(sorted_folder, f"sorted_{timestamp}_{input_filename}")

        print(f"\nReading file: {input_file}")
        
        # Detect and skip header rows
        skip_rows = detect_header_row(input_file)
        
        # Read the Excel file
        df = pd.read_excel(input_file, skiprows=skip_rows)
        print(f"File read successfully. Found {len(df)} rows.")
        
        # Print available columns
        print("\nColumns found in your file:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")

        # Find matching column (case-insensitive)
        column_found = None
        for col in df.columns:
            if isinstance(col, str) and col.lower() == sort_by_column.lower():
                column_found = col
                break

        if column_found:
            print(f"\nSorting by: {column_found}")
            
            # Handle date columns
            if 'date' in column_found.lower():
                df[column_found] = pd.to_datetime(df[column_found], errors='coerce')
            
            # Sort the dataframe
            df_sorted = df.sort_values(by=column_found, ascending=True, na_position='last')
            
            # Save the sorted file
            print(f"\nSaving sorted file to: {output_file}")
            df_sorted.to_excel(output_file, index=False)
            
            if os.path.exists(output_file):
                file_size = os.path.getsize(output_file)
                print(f"File saved successfully! Size: {file_size:,} bytes")
                print("\nFirst few rows of sorted data:")
                print(df_sorted[list(df_sorted.columns)[:5]].head())  # Show first 5 columns
                print(f"\nTotal rows sorted: {len(df_sorted):,}")
            else:
                print("Error: File was not created!")
        else:
            print(f"\nError: Column '{sort_by_column}' not found in file.")
            print("Available columns are:")
            for col in df.columns:
                print(f"- {col}")
            
    except Exception as e:
        print(f"\nError processing file: {str(e)}")
        print("Please check if:")
        print("1. The Excel file exists and isn't open")
        print("2. You have permission to write to the output folder")
        print("3. The column name matches exactly")

def print_usage_guide():
    """Print guide for using the script"""
    print("\nUsage Guide:")
    print("1. Make sure your Excel file is closed")
    print("2. The script will show you available columns")
    print("3. Column names are case-insensitive")
    print("4. Sorted files are saved with timestamps\n")

# Example usage
if __name__ == "__main__":
    # Settings - modify these as needed
    input_filename = "ITAD Recieving.xlsx"  # Your Excel file name
    sort_column = "Item Name"               # Column to sort by

    print_usage_guide()
    sort_excel_file(input_filename, sort_column)