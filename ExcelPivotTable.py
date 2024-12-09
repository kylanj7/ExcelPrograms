import pandas as pd

def create_pivot_table(filename, index_cols, value_cols, agg_func='sum'):
    """Create a pivot table from Excel data"""
    df = pd.read_excel(filename)
    pivot = pd.pivot_table(df, 
                          index=index_cols,
                          values=value_cols,
                          aggfunc=agg_func)
    pivot.to_excel(f'pivot_{filename}')

# Example usage
if __name__ == "__main__":
    data = {
        'Department': ['Sales', 'IT', 'Sales', 'IT'],
        'Revenue': [100, 200, 150, 250]
    }
    df = pd.DataFrame(data)
    df.to_excel('sales.xlsx', index=False)
    create_pivot_table('sales.xlsx', 
                      index_cols=['Department'],
                      value_cols=['Revenue'])