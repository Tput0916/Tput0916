import pandas as pd
import os

def convert_excel_to_csv():
    """
    Converts the 'refer.xlsx' file to 'ground_truth.csv' in the same directory.
    """
    # Define the absolute paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, 'refer.xlsx')
    csv_path = os.path.join(script_dir, 'ground_truth.csv')

    print(f"Reading from: {excel_path}")
    
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
        return

    try:
        # Read the Excel file
        df = pd.read_excel(excel_path)
        
        # Write to a CSV file
        df.to_csv(csv_path, index=False)
        
        print(f"Successfully converted to: {csv_path}")

    except Exception as e:
        print(f"An error occurred during conversion: {e}")

if __name__ == '__main__':
    convert_excel_to_csv()