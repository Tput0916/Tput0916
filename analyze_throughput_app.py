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


################
def main():
    """
    主函数，用于运行分析。
    """
    print("--- 自动化吞吐率分析工具 ---")
    print("\n重要提示：如果需要与Excel文件进行比较，请确保在运行前关闭该文件。")
    
    # 获取用户拖拽的文件夹路径
    target_folder_input = input("请将包含数据的数据文件夹拖到此处，然后按Enter键：")
    # 清理路径字符串，这对于处理Windows和macOS拖拽的路径特别重要
    target_folder = target_folder_input.strip().strip('\'"').strip()

    if not os.path.isdir(target_folder):
        print(f"\n错误：提供的路径 '{target_folder}' 不是一个有效的文件夹。")
        input("按Enter键退出。") # 让窗口保持打开，以便用户能看到错误信息
        return

    # *** 修改部分 ***
    # 在用户提供的数据文件夹内部创建输出目录
    output_dir = os.path.join(target_folder, 'SportonAutoData_Output')
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n--- 输出文件将保存到: {output_dir} ---")

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    script_df_long = extract_script_data(base_dir=target_folder)

    if script_df_long is None or script_df_long.empty:
        print("\n处理已停止：未能从CSV文件中提取到任何数据。")
        input("按Enter键退出。")
        return

    # --- 函数的其余部分保持不变 ---
    # ... (包括汇总、计算下降率、生成CSV报告和图表) ...
    
    # (这里是你 main 函数中其余未变动的代码)

    # --- 根据条件执行比较 ---
    excel_files = [f for f in os.listdir(target_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    if len(excel_files) == 1:
        manual_data_path = os.path.join(target_folder, excel_files[0])
        print(f"\n--- 已找到手动数据文件: '{excel_files[0]}'。开始进行比较。 ---")
        perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
    elif len(excel_files) > 1:
        manual_data_path = os.path.join(target_folder, excel_files[0])
        print(f"\n--- 警告: 发现多个Excel文件。将使用找到的第一个文件: '{excel_files[0]}'。 ---")
        perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
    else:
        print(f"\n--- 未找到手动数据的Excel文件。跳过比较步骤。 ---")

    print("\n--- 分析成功完成！ ---")
    input("按Enter键退出。") # 让窗口保持打开，以便用户能看到完成信息
    



if __name__ == '__main__':
    main()