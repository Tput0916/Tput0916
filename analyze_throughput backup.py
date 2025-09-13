import os
import pandas as pd
import matplotlib.pyplot as plt
import re
from datetime import datetime

def extract_script_data(base_dir):
    """
    Extracts throughput data from CSV files in subdirectories of the provided base_dir.
    """
    subdirectories = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d)) and not d.startswith('.')]
    all_data = []

    print("--- Scanning subdirectories for CSV files... ---")
    for subdir in subdirectories:
        subdir_path = os.path.join(base_dir, subdir)
        for filename in os.listdir(subdir_path):
            if filename.endswith('.csv'):
                match = re.match(r'(\d+)-', filename)
                if match:
                    rssi = -int(match.group(1))
                    filepath = os.path.join(subdir_path, filename)
                    try:
                        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                            lines = f.readlines()
                        
                        throughput = None
                        for i in range(len(lines) - 1, -1, -1):
                            if 'Throughput Avg.(Mbps)' in lines[i]:
                                header_line_index = i
                                header_fields = [field.strip().strip('"') for field in lines[i].split(',')]
                                try:
                                    throughput_col_index = header_fields.index('Throughput Avg.(Mbps)')
                                    if header_line_index + 1 < len(lines):
                                        data_line = lines[header_line_index + 1]
                                        data_fields = data_line.split(',')
                                        if throughput_col_index < len(data_fields):
                                            throughput_str = data_fields[throughput_col_index].strip()
                                            if throughput_str:
                                                throughput = float(throughput_str)
                                                break
                                except ValueError:
                                    continue
                        
                        if throughput is not None:
                            all_data.append({'folder': subdir, 'rssi': rssi, 'throughput': throughput})
                    except Exception as e:
                        print(f"Warning: Could not process file {filepath}: {e}")
    
    if not all_data:
        return None
        
    return pd.DataFrame(all_data)

def perform_comparison(script_df_long, manual_data_path, output_dir, timestamp):
    """
    Compares script-generated data with manual data from the Excel file.
    """
    print("--- Reading manual data from Excel file for comparison... ---")
    sheets_to_compare = script_df_long['folder'].unique()
    all_comparison_data = []

    for sheet in sheets_to_compare:
        try:
            manual_sheet_df_full = pd.read_excel(manual_data_path, sheet_name=sheet, header=None)
            
            rssi_coords, throughput_coords = None, None
            for r_idx, row in manual_sheet_df_full.iterrows():
                for c_idx, cell in enumerate(row):
                    cell_str = str(cell)
                    if 'AP-RSSI (dBm)' in cell_str:
                        rssi_coords = (r_idx, c_idx)
                    elif sheet in cell_str:
                        throughput_coords = (r_idx, c_idx)
                if rssi_coords and throughput_coords:
                    break

            if rssi_coords and throughput_coords:
                rssi_row_data = manual_sheet_df_full.iloc[rssi_coords[0], rssi_coords[1] + 1:]
                throughput_row_data = manual_sheet_df_full.iloc[throughput_coords[0], throughput_coords[1] + 1:]

                temp_df = pd.DataFrame([rssi_row_data.values, throughput_row_data.values]).T
                temp_df.columns = ['rssi', 'Manual Throughput']
                temp_df['rssi'] = pd.to_numeric(temp_df['rssi'], errors='coerce')
                temp_df['Manual Throughput'] = pd.to_numeric(temp_df['Manual Throughput'], errors='coerce')
                manual_sheet_df = temp_df.dropna().set_index('rssi')

                script_series = script_df_long[script_df_long['folder'] == sheet].set_index('rssi')[['throughput']].rename(columns={'throughput': 'Script Throughput'})
                
                comparison_df = pd.merge(script_series, manual_sheet_df, on='rssi', how='outer')
                comparison_df['Delta'] = (comparison_df['Script Throughput'] - comparison_df['Manual Throughput']).round(1)
                comparison_df['Scenario'] = sheet
                all_comparison_data.append(comparison_df.reset_index())
            else:
                print(f"Warning: Could not find required headers in Excel sheet '{sheet}'.")
        except Exception as e:
            print(f"Warning: Could not process Excel sheet '{sheet}': {e}")

    if not all_comparison_data:
        print("Error: No comparison data was generated. Check Excel file structure.")
        return

    final_comparison_df = pd.concat(all_comparison_data)
    summary_pivot = final_comparison_df.pivot_table(index='rssi', columns='Scenario', values=['Script Throughput', 'Manual Throughput', 'Delta'])
    summary_pivot.sort_index(ascending=False, inplace=True)
    summary_output_path = os.path.join(output_dir, f"{timestamp}_comparison_summary.csv")
    summary_pivot.to_csv(summary_output_path)
    print(f"Comparison summary saved to {summary_output_path}")

    scenarios = final_comparison_df['Scenario'].unique()
    num_cols = 2
    num_rows = (len(scenarios) + 1) // num_cols
    fig, axes = plt.subplots(num_rows, num_cols, figsize=(15, 5 * num_rows), squeeze=False)
    axes = axes.flatten()

    for i, scenario in enumerate(scenarios):
        ax = axes[i]
        scenario_df = final_comparison_df[final_comparison_df['Scenario'] == scenario].set_index('rssi').sort_index(ascending=False)
        ax.plot(scenario_df.index, scenario_df['Script Throughput'], marker='o', linestyle='-', label='Script Data')
        ax.plot(scenario_df.index, scenario_df['Manual Throughput'], marker='x', linestyle='--', label='Manual Data')
        ax.set_title(f'Comparison for {scenario}')
        ax.set_xlabel('RSSI (dBm)')
        ax.set_ylabel('Throughput (Mbps)')
        ax.legend()
        ax.grid(True)
        ax.invert_xaxis()

    for j in range(i + 1, len(axes)):
        fig.delaxes(axes[j])

    plt.tight_layout()
    plot_output_path = os.path.join(output_dir, f"{timestamp}_comparison_plot.png")
    plt.savefig(plot_output_path)
    print(f"Comparison plot saved to {plot_output_path}")
    plt.close()

def main():
    """
    Main function to run the analysis.
    """
    print("--- Automated Throughput Analysis Tool ---")
    
    # 1. Prompt for input path
    target_folder = input("Please drag the folder containing your data here and press Enter: ").strip()
    target_folder = target_folder.strip('\'"')

    if not os.path.isdir(target_folder):
        print(f"\nError: The provided path '{target_folder}' is not a valid directory.")
        return

    # 2. Define output directory relative to the script's location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'SportonAutoData')
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n--- Output will be saved to: {output_dir} ---")

    # 3. Generate a single timestamp for all output files from this run
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # 4. Run core analysis
    script_df_long = extract_script_data(base_dir=target_folder)

    if script_df_long is None or script_df_long.empty:
        print("\nProcessing stopped: No data was extracted from CSV files in the provided folder.")
        return

    # --- Generate and save the primary analysis files ---
    summary_df = script_df_long.pivot_table(index='rssi', columns='folder', values='throughput')
    summary_df.sort_index(ascending=False, inplace=True)
    summary_path = os.path.join(output_dir, f"{timestamp}_throughput_summary.csv")
    summary_df.to_csv(summary_path)
    print(f"Primary summary table saved to {summary_path}")

    plt.figure(figsize=(12, 8))
    for folder_name, group in script_df_long.groupby('folder'):
        group = group.sort_values(by='rssi', ascending=False)
        plt.plot(group['rssi'], group['throughput'], marker='o', linestyle='-', label=folder_name)
    plt.title('Throughput vs. RSSI')
    plt.xlabel('RSSI (dBm)')
    plt.ylabel('Throughput (Mbps)')
    plt.legend()
    plt.grid(True)
    plt.gca().invert_xaxis()
    plot_path = os.path.join(output_dir, f"{timestamp}_throughput_analysis.png")
    plt.savefig(plot_path)
    plt.close()
    print(f"Primary analysis plot saved to {plot_path}")
    
    # --- Conditionally perform comparison ---
    manual_data_path = os.path.join(target_folder, 'Huawei MateBook X Pro_WiFi Conductive Throughput_Testing Data_20250909.xlsx')
    if os.path.exists(manual_data_path):
        print(f"\n--- Found manual data file. Starting comparison. ---")
        perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
        print("--- Comparison complete. ---")
    else:
        print(f"\n--- Manual data file not found in the provided folder. Skipping comparison. ---")

    print("\n--- Analysis finished successfully! ---")

if __name__ == '__main__':
    main()