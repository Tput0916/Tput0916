import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
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
    print("\n--- Reading manual data from Excel file for comparison... ---")
    sheets_to_compare = script_df_long['folder'].unique()
    all_comparison_data = []
    for sheet in sheets_to_compare:
        try:
            manual_sheet_df_full = pd.read_excel(manual_data_path, sheet_name=sheet, header=None)
            rssi_coords, throughput_coords = None, None
            for r_idx, row in manual_sheet_df_full.iterrows():
                for c_idx, cell in enumerate(row):
                    cell_str = str(cell)
                    if 'AP-RSSI (dBm)' in cell_str: rssi_coords = (r_idx, c_idx)
                    elif sheet in cell_str: throughput_coords = (r_idx, c_idx)
                if rssi_coords and throughput_coords: break
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
            else: print(f"Warning: Could not find required headers in Excel sheet '{sheet}'.")
        except Exception as e: print(f"Warning: Could not process Excel sheet '{sheet}': {e}")
    if not all_comparison_data:
        print("Error: No comparison data was generated.")
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
    for j in range(i + 1, len(axes)): fig.delaxes(axes[j])
    plt.tight_layout()
    plot_output_path = os.path.join(output_dir, f"{timestamp}_comparison_plot.png")
    plt.savefig(plot_output_path)
    plt.close()
    print(f"Comparison plot saved to {plot_output_path}")

def main():
    """
    Main function to run the analysis.
    """
    print("--- Automated Throughput Analysis Tool ---")
    print("\nIMPORTANT: If comparing with an Excel file, please ensure it is CLOSED before proceeding.")
    
    # Use a graphical file dialog for a better user experience
    from tkinter import filedialog, Tk
    print("--- Please select the folder containing your data in the dialog window... ---")
    root = Tk()
    root.withdraw() # Hide the main tkinter window
    target_folder = filedialog.askdirectory(title="Please select the folder containing your data")
    root.destroy()

    if not target_folder: # If the user closes the dialog
        print("\nNo folder selected. Exiting.")
        return

    if not os.path.isdir(target_folder):
        print(f"\nError: The provided path '{target_folder}' is not a valid directory.")
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'SportonAutoData')
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n--- Output will be saved to: {output_dir} ---")

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    script_df_long = extract_script_data(base_dir=target_folder)

    if script_df_long is None or script_df_long.empty:
        print("\nProcessing stopped: No data was extracted from CSV files.")
        return

    # --- Create the main throughput summary table ---
    summary_df = script_df_long.pivot_table(index='rssi', columns='folder', values='throughput')
    summary_df.sort_index(ascending=False, inplace=True)

    # --- Calculate Drop Rate and prepare for combined report ---
    drop_rate_df = pd.DataFrame(index=summary_df.index)
    if 'Base Line' in summary_df.columns:
        print("\n--- 'Base Line' found. Calculating Drop Rate... ---")
        baseline_throughput = summary_df['Base Line']
        non_baseline_scenarios = [col for col in summary_df.columns if col != 'Base Line']
        
        for scenario in non_baseline_scenarios:
            non_baseline_throughput = summary_df[scenario]
            drop_rate = (baseline_throughput - non_baseline_throughput).divide(baseline_throughput).replace([float('inf'), -float('inf')], 0)
            drop_rate_df[scenario] = drop_rate
    else:
        print("\n--- 'Base Line' data not found. Skipping Drop Rate calculation. ---")

    # --- Create the single, integrated CSV report ---
    throughput_cols = pd.MultiIndex.from_product([['Throughput (Mbps)'], summary_df.columns])
    drop_rate_cols = pd.MultiIndex.from_product([['Drop Rate (%)'], drop_rate_df.columns])
    
    full_report_df = pd.DataFrame(summary_df.values, index=summary_df.index, columns=throughput_cols)
    drop_rate_report_df = pd.DataFrame((drop_rate_df * 100).round(1).values, index=drop_rate_df.index, columns=drop_rate_cols)
    
    final_report_df = pd.concat([full_report_df, drop_rate_report_df], axis=1)
    
    report_path_csv = os.path.join(output_dir, f"{timestamp}_full_report.csv")
    final_report_df.to_csv(report_path_csv)
    print(f"Full report saved to {report_path_csv}")

    # --- Create the single, integrated Image report ---
    print("--- Generating combined chart... ---")
    
    prop_cycle = plt.rcParams['axes.prop_cycle']
    colors = prop_cycle.by_key()['color']
    color_map = {scenario: colors[i % len(colors)] for i, scenario in enumerate(summary_df.columns)}

    fig, axes = plt.subplots(2, 1, figsize=(20, 18))
    
    ax_tp = axes[0]
    for scenario in summary_df.columns:
        ax_tp.plot(summary_df.index, summary_df[scenario], marker='o', linestyle='-', label=scenario, color=color_map[scenario])
    ax_tp.set_title('Throughput Analysis', fontsize=22)
    ax_tp.set_ylabel('Throughput (Mbps)', fontsize=16)
    ax_tp.legend(fontsize=14)
    ax_tp.grid(True)
    plt.setp(ax_tp.get_xticklabels(), visible=True)
    ax_tp.tick_params(axis='both', which='major', labelsize=14)

    ax_dr = axes[1]
    if not drop_rate_df.empty:
        for scenario in drop_rate_df.columns:
            ax_dr.plot(drop_rate_df.index, drop_rate_df[scenario], marker='o', linestyle='-', label=scenario, color=color_map[scenario])
        ax_dr.yaxis.set_major_formatter(mtick.PercentFormatter(1.0))
        ax_dr.set_ylabel('Drop Rate', fontsize=16)
        ax_dr.legend(fontsize=14)
        ax_dr.grid(True)
        ax_dr.tick_params(axis='both', which='major', labelsize=14)
        ax_dr.set_title('Drop Rate Analysis', fontsize=22)
    
    ax_dr.set_xlabel('RSSI (dBm)', fontsize=16)
    ax_tp.invert_xaxis()
    ax_dr.invert_xaxis()
    
    plt.subplots_adjust(left=0.15, right=0.9, top=0.9, bottom=0.1)
    
    report_path_png = os.path.join(output_dir, f"{timestamp}_combined_chart.png")
    plt.savefig(report_path_png, bbox_inches='tight')
    plt.close()
    print(f"Combined chart saved to {report_path_png}")

    # --- Conditionally perform comparison ---
    # Find the manual data Excel file automatically
    excel_files = [f for f in os.listdir(target_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    if len(excel_files) == 1:
        manual_data_path = os.path.join(target_folder, excel_files[0])
        print(f"\n--- Found manual data file: '{excel_files[0]}'. Starting comparison. ---")
        perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
    elif len(excel_files) > 1:
        manual_data_path = os.path.join(target_folder, excel_files[0])
        print(f"\n--- Warning: Found multiple Excel files. Using the first one found: '{excel_files[0]}'. ---")
        perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
    else:
        print(f"\n--- No manual data Excel file found. Skipping comparison. ---")

    print("\n--- Analysis finished successfully! ---")

if __name__ == '__main__':
    main()