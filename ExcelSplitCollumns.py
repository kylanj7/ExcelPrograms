import pandas as pd

def split_excel_by_column(filename, column):
    """Split Excel file into multiple files based on unique values in a column"""
    df = pd.read_excel(filename)
    for value in df[column].unique():
        df_split = df[df[column] == value]
        df_split.to_excel(f'{value}_{filename}', index=False)

# Example usage
if __name__ == "__main__":
    data = {
        'Department': ['Sales', 'IT', 'Sales', 'IT'],
        'Name': ['Alice', 'Bob', 'Carol', 'Dave']
    }
    df = pd.DataFrame(data)
    df.to_excel('employees.xlsx', index=False)
    split_excel_by_column('employees.xlsx', 'Department')