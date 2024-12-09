import pandas as pd

def remove_duplicates(filename, columns):
    """Remove duplicate rows based on specified columns"""
    df = pd.read_excel(filename)
    df_clean = df.drop_duplicates(subset=columns)
    df_clean.to_excel(f'cleaned_{filename}', index=False)

# Example usage
if __name__ == "__main__":
    data = {
        'ID': [1, 1, 2],
        'Name': ['Alice', 'Alice', 'Bob']
    }
    df = pd.DataFrame(data)
    df.to_excel('duplicates.xlsx', index=False)
    remove_duplicates('duplicates.xlsx', ['ID', 'Name'])