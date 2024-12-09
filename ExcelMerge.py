import pandas as pd

def merge_excel_files(file_list, output_name):
    """Merge multiple Excel files into one"""
    combined = pd.DataFrame()
    for file in file_list:
        df = pd.read_excel(file)
        combined = pd.concat([combined, df])
    combined.to_excel(output_name, index=False)

# Example usage
if __name__ == "__main__":
    # Create sample files
    df1 = pd.DataFrame({'ID': [1, 2], 'Name': ['Alice', 'Bob']})
    df2 = pd.DataFrame({'ID': [3, 4], 'Name': ['Carol', 'Dave']})
    
    df1.to_excel('file1.xlsx', index=False)
    df2.to_excel('file2.xlsx', index=False)
    
    merge_excel_files(['file1.xlsx', 'file2.xlsx'], 'merged.xlsx')