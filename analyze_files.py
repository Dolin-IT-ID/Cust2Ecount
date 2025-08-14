import pandas as pd
import os

def analyze_excel_structure(file_path):
    """Analyze Excel file structure and return column information"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        
        print(f"\n=== Analysis of {os.path.basename(file_path)} ===")
        print(f"Shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
        print(f"Data types:")
        print(df.dtypes)
        print(f"\nFirst 5 rows:")
        print(df.head())
        print(f"\nColumn info:")
        print(df.info())
        
        return {
            'columns': list(df.columns),
            'shape': df.shape,
            'dtypes': df.dtypes.to_dict(),
            'sample_data': df.head().to_dict()
        }
        
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")
        return None

if __name__ == "__main__":
    # Analyze both files
    template_file = "Template_Ecount.xlsx"
    source_file = "SCRAP_DATA.xlsx"
    
    print("Analyzing Excel files structure...")
    
    # Analyze template file
    template_info = analyze_excel_structure(template_file)
    
    # Analyze source file
    source_info = analyze_excel_structure(source_file)
    
    # Compare columns
    if template_info and source_info:
        print("\n=== COLUMN COMPARISON ===")
        print(f"Template columns: {template_info['columns']}")
        print(f"Source columns: {source_info['columns']}")
        
        # Find potential matches (case-insensitive)
        template_cols = [col.lower() for col in template_info['columns']]
        source_cols = [col.lower() for col in source_info['columns']]
        
        matches = []
        for t_col in template_info['columns']:
            for s_col in source_info['columns']:
                if t_col.lower() == s_col.lower():
                    matches.append((t_col, s_col))
        
        print(f"\nExact matches found: {matches}")
        
        # Find similar columns (contains substring)
        similar = []
        for t_col in template_info['columns']:
            for s_col in source_info['columns']:
                if (t_col.lower() in s_col.lower() or s_col.lower() in t_col.lower()) and t_col.lower() != s_col.lower():
                    similar.append((t_col, s_col))
        
        print(f"Similar columns found: {similar}")