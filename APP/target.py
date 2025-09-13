# target.py (v3.1 - Final Professional Structure)

# 关键优化1：强制Matplotlib使用Agg后端
import matplotlib
matplotlib.use('Agg')

import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import re
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image
import threading
# 关键优化2：导入拖拽库
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- 核心数据处理函数 (这部分代码与之前完全相同，此处为简洁省略) ---
# ... (请将你之前的 extract_script_data 和 perform_comparison 函数完整地复制到这里) ...
def extract_script_data(base_dir):
    # (此函数代码与之前完全相同，此处为简洁省略)
    subdirectories = [d for d in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, d)) and not d.startswith('.')]
    all_data = []
    for subdir in subdirectories:
        subdir_path = os.path.join(base_dir, subdir)
        for filename in os.listdir(subdir_path):
            if filename.endswith('.csv'):
                match = re.match(r'(\d+)-', filename)
                if match:
                    rssi = -int(match.group(1))
                    filepath = os.path.join(subdir_path, filename)
                    try:
                        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f: lines = f.readlines()
                        throughput = None
                        for i in range(len(lines) - 1, -1, -1):
                            if 'Throughput Avg.(Mbps)' in lines[i]:
                                header_line_index = i
                                header_fields = [field.strip().strip('"') for field in lines[i].split(',')]
                                try:
                                    throughput_col_index = header_fields.index('Throughput Avg.(Mbps)')
                                    if header_line_index + 1 < len(lines):
                                        data_line = lines[header_line_index + 1]; data_fields = data_line.split(',')
                                        if throughput_col_index < len(data_fields):
                                            throughput_str = data_fields[throughput_col_index].strip()
                                            if throughput_str: throughput = float(throughput_str); break
                                except ValueError: continue
                        if throughput is not None: all_data.append({'folder': subdir, 'rssi': rssi, 'throughput': throughput})
                    except Exception as e: print(f"Warning: Could not process file {filepath}: {e}")
    if not all_data: return None
    return pd.DataFrame(all_data)

def perform_comparison(script_df_long, manual_data_path, output_dir, timestamp):
    # (此函数代码与之前完全相同，此处为简洁省略)
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
                temp_df.columns = ['rssi', 'Manual Throughput']; temp_df['rssi'] = pd.to_numeric(temp_df['rssi'], errors='coerce'); temp_df['Manual Throughput'] = pd.to_numeric(temp_df['Manual Throughput'], errors='coerce')
                manual_sheet_df = temp_df.dropna().set_index('rssi')
                script_series = script_df_long[script_df_long['folder'] == sheet].set_index('rssi')[['throughput']].rename(columns={'throughput': 'Script Throughput'})
                comparison_df = pd.merge(script_series, manual_sheet_df, on='rssi', how='outer')
                comparison_df['Delta'] = (comparison_df['Script Throughput'] - comparison_df['Manual Throughput']).round(1); comparison_df['Scenario'] = sheet
                all_comparison_data.append(comparison_df.reset_index())
        except Exception as e: print(f"Warning: Could not process Excel sheet '{sheet}': {e}")
    if not all_comparison_data: return None
    final_comparison_df = pd.concat(all_comparison_data)
    summary_pivot = final_comparison_df.pivot_table(index='rssi', columns='Scenario', values=['Script Throughput', 'Manual Throughput', 'Delta'])
    summary_pivot.sort_index(ascending=False, inplace=True)
    summary_output_path = os.path.join(output_dir, f"{timestamp}_comparison_summary.csv")
    summary_pivot.to_csv(summary_output_path)
    scenarios = final_comparison_df['Scenario'].unique()
    num_cols = 2; num_rows = (len(scenarios) + 1) // num_cols
    fig, axes = plt.subplots(num_rows, num_cols, figsize=(15, 5 * num_rows), squeeze=False); axes = axes.flatten()
    for i, scenario in enumerate(scenarios):
        ax = axes[i]; scenario_df = final_comparison_df[final_comparison_df['Scenario'] == scenario].set_index('rssi').sort_index(ascending=False)
        ax.plot(scenario_df.index, scenario_df['Script Throughput'], marker='o', linestyle='-', label='Script Data')
        ax.plot(scenario_df.index, scenario_df['Manual Throughput'], marker='x', linestyle='--', label='Manual Data')
        ax.set_title(f'Comparison for {scenario}'); ax.set_xlabel('RSSI (dBm)'); ax.set_ylabel('Throughput (Mbps)'); ax.legend(); ax.grid(True); ax.invert_xaxis()
    for j in range(i + 1, len(axes)): fig.delaxes(axes[j])
    plt.tight_layout()
    plot_output_path = os.path.join(output_dir, f"{timestamp}_comparison_plot.png")
    plt.savefig(plot_output_path); plt.close()
    return plot_output_path


# 关键优化：App现在是内容框架，而不是整个窗口
class AppFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.grid_columnconfigure(0, weight=1)

        # ---- UI 元素 ----
        self.path_frame = ctk.CTkFrame(self)
        self.path_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.path_frame.grid_columnconfigure(1, weight=1)
        self.entry_path = ctk.CTkEntry(self.path_frame, placeholder_text="Drag & drop a folder here, or click 'Browse...'")
        self.entry_path.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        # 注册拖拽事件
        self.entry_path.drop_target_register(DND_FILES)
        self.entry_path.dnd_bind('<<Drop>>', self.on_drop)
        
        # ... (其余UI元素与之前相同) ...
        self.label_path = ctk.CTkLabel(self.path_frame, text="Raw Data Path:", font=ctk.CTkFont(size=14, weight="bold"))
        self.label_path.grid(row=0, column=0, padx=10, pady=10)
        self.button_browse = ctk.CTkButton(self.path_frame, text="Browse...", command=self.select_folder)
        self.button_browse.grid(row=0, column=2, padx=10, pady=10)
        self.control_frame = ctk.CTkFrame(self)
        self.control_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.control_frame.grid_columnconfigure((0, 1), weight=1)
        self.button_cancel = ctk.CTkButton(self.control_frame, text="Quit", command=self.master.quit) # 注意：命令是 self.master.quit
        self.button_cancel.grid(row=0, column=0, padx=20, pady=10, sticky="e")
        self.button_ok = ctk.CTkButton(self.control_frame, text="Start Analysis", command=self.start_analysis)
        self.button_ok.grid(row=0, column=1, padx=20, pady=10, sticky="w")
        self.result_frame = ctk.CTkFrame(self)
        self.result_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.grid_rowconfigure(2, weight=1) # 让结果区域可扩展
        self.result_frame.grid_columnconfigure(0, weight=1)
        self.result_frame.grid_rowconfigure(2, weight=1) 
        self.label_result_path = ctk.CTkLabel(self.result_frame, text="Welcome! Please select a data folder and click 'Start Analysis'.", text_color="gray", wraplength=700)
        self.label_result_path.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.progress_bar = ctk.CTkProgressBar(self.result_frame, orientation="horizontal", mode="indeterminate")
        self.progress_bar.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        self.progress_bar.grid_remove() 
        self.scrollable_frame = ctk.CTkScrollableFrame(self.result_frame, label_text="Result Images")
        self.scrollable_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

    def on_drop(self, event):
        path = event.data.strip('{}')
        self.entry_path.delete(0, "end")
        self.entry_path.insert(0, path)

    # --- 后面的所有分析和UI更新函数都与上一版相同 ---
    # ... (请将你之前的 _run_analysis_thread, select_folder, start_analysis 等函数完整地复制到这里) ...
    def _run_analysis_thread(self, target_folder):
        try:
            output_dir = os.path.join(target_folder, 'SportonAutoData_Output')
            os.makedirs(output_dir, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            script_df_long = extract_script_data(base_dir=target_folder)
            if script_df_long is None or script_df_long.empty: raise ValueError("No data was extracted from CSV files.")
            summary_df = script_df_long.pivot_table(index='rssi', columns='folder', values='throughput')
            summary_df.sort_index(ascending=False, inplace=True)
            drop_rate_df = pd.DataFrame(index=summary_df.index)
            if 'Base Line' in summary_df.columns:
                baseline_throughput = summary_df['Base Line']
                non_baseline_scenarios = [col for col in summary_df.columns if col != 'Base Line']
                for scenario in non_baseline_scenarios:
                    drop_rate = (baseline_throughput - summary_df[scenario]).divide(baseline_throughput).replace([float('inf'), -float('inf')], 0)
                    drop_rate_df[scenario] = drop_rate
            throughput_cols = pd.MultiIndex.from_product([['Throughput (Mbps)'], summary_df.columns])
            drop_rate_cols = pd.MultiIndex.from_product([['Drop Rate (%)'], drop_rate_df.columns])
            full_report_df = pd.DataFrame(summary_df.values, index=summary_df.index, columns=throughput_cols)
            drop_rate_report_df = pd.DataFrame((drop_rate_df * 100).round(1).values, index=drop_rate_df.index, columns=drop_rate_cols)
            final_report_df = pd.concat([full_report_df, drop_rate_report_df], axis=1)
            report_path_csv = os.path.join(output_dir, f"{timestamp}_full_report.csv")
            final_report_df.to_csv(report_path_csv)
            print(f"Full report saved to {report_path_csv}")
            prop_cycle = plt.rcParams['axes.prop_cycle']
            colors = prop_cycle.by_key()['color']
            color_map = {scenario: colors[i % len(colors)] for i, scenario in enumerate(summary_df.columns)}
            fig, axes = plt.subplots(2, 1, figsize=(20, 18))
            ax_tp = axes[0]
            for scenario in summary_df.columns: ax_tp.plot(summary_df.index, summary_df[scenario], marker='o', linestyle='-', label=scenario, color=color_map[scenario])
            ax_tp.set_title('Throughput Analysis', fontsize=22); ax_tp.set_ylabel('Throughput (Mbps)', fontsize=16); ax_tp.legend(fontsize=14); ax_tp.grid(True)
            plt.setp(ax_tp.get_xticklabels(), visible=True); ax_tp.tick_params(axis='both', which='major', labelsize=14)
            ax_dr = axes[1]
            if not drop_rate_df.empty:
                for scenario in drop_rate_df.columns: ax_dr.plot(drop_rate_df.index, drop_rate_df[scenario], marker='o', linestyle='-', label=scenario, color=color_map[scenario])
                ax_dr.yaxis.set_major_formatter(mtick.PercentFormatter(1.0)); ax_dr.set_ylabel('Drop Rate', fontsize=16); ax_dr.legend(fontsize=14); ax_dr.grid(True)
                ax_dr.tick_params(axis='both', which='major', labelsize=14); ax_dr.set_title('Drop Rate Analysis', fontsize=22)
            ax_dr.set_xlabel('RSSI (dBm)', fontsize=16); ax_tp.invert_xaxis(); ax_dr.invert_xaxis()
            plt.subplots_adjust(left=0.15, right=0.9, top=0.9, bottom=0.1)
            report_path_png_combined = os.path.join(output_dir, f"{timestamp}_combined_chart.png")
            plt.savefig(report_path_png_combined, bbox_inches='tight'); plt.close(fig)
            print(f"Combined chart saved to {report_path_png_combined}")
            comparison_plot_path = None
            excel_files = [f for f in os.listdir(target_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
            if excel_files:
                manual_data_path = os.path.join(target_folder, excel_files[0])
                comparison_plot_path = perform_comparison(script_df_long, manual_data_path, output_dir, timestamp)
            self.after(0, self.analysis_finished, output_dir, report_path_png_combined, comparison_plot_path)
        except Exception as e:
            self.after(0, self.analysis_failed, str(e))
    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: self.entry_path.delete(0, "end"); self.entry_path.insert(0, folder_path)
    def start_analysis(self):
        target_folder = self.entry_path.get()
        if not target_folder or not os.path.isdir(target_folder): self.label_result_path.configure(text="Error: Please select a valid folder first!", text_color="red"); return
        self.set_ui_state("disabled")
        for widget in self.scrollable_frame.winfo_children(): widget.destroy()
        analysis_thread = threading.Thread(target=self._run_analysis_thread, args=(target_folder,)); analysis_thread.start()
    def set_ui_state(self, state: str):
        self.button_ok.configure(state=state); self.button_browse.configure(state=state); self.entry_path.configure(state=state)
        if state == "disabled": self.progress_bar.grid(); self.progress_bar.start(); self.label_result_path.configure(text="Analyzing data, please wait...", text_color="gray")
        else: self.progress_bar.stop(); self.progress_bar.grid_remove()
    def analysis_finished(self, output_dir, img_path1, img_path2):
        self.set_ui_state("normal"); self.label_result_path.configure(text=f"Success! All results saved in: {output_dir}", text_color="green")
        self.create_and_show_image(img_path1); 
        if img_path2: self.create_and_show_image(img_path2)
    def analysis_failed(self, error_message):
        self.set_ui_state("normal"); self.label_result_path.configure(text=f"Error: {error_message}", text_color="red")
    def create_and_show_image(self, path, max_width=700):
        try:
            image = Image.open(path); ratio = max_width / image.width; new_height = int(image.height * ratio)
            ctk_image = ctk.CTkImage(light_image=image, size=(max_width, new_height))
            image_label = ctk.CTkLabel(self.scrollable_frame, image=ctk_image, text=""); image_label.pack(padx=10, pady=10, fill="x", expand=True)
        except Exception as e:
            error_label = ctk.CTkLabel(self.scrollable_frame, text=f"Could not load image:\n{os.path.basename(path)}\nError: {e}"); error_label.pack(padx=10, pady=10)


if __name__ == "__main__":
    # 使用CTK设置外观
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    # 关键优化：创建DND功能的根窗口
    root = TkinterDnD.Tk()
    root.title("Throughput Analyzer")
    root.geometry("800x800")
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    # 将我们的内容框架嵌入根窗口
    app_frame = AppFrame(master=root)
    app_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

    root.mainloop()
