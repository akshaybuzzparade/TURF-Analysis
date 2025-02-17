import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from itertools import combinations as comb

class TurfAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TURF Analysis Application")
        self.root.geometry("1200x750")

        # Top Frame for File Upload and Subset Selection
        self.top_frame = tk.Frame(root)
        self.top_frame.pack(pady=10, fill="x")

        self.upload_button = tk.Button(self.top_frame, text="Upload Excel File", command=self.load_file)
        self.upload_button.pack(side="left", padx=10)

        tk.Label(self.top_frame, text="Subset Size:").pack(side="left")
        self.subset_size_entry = tk.Entry(self.top_frame, width=5)
        self.subset_size_entry.pack(side="left", padx=5)

        self.run_analysis_button = tk.Button(self.top_frame, text="Run TURF Analysis", command=self.run_turf_analysis)
        self.run_analysis_button.pack(side="left", padx=10)

        # Filters Frame
        self.filters_frame = tk.LabelFrame(root, text="Filters", padx=10, pady=10)
        self.filters_frame.pack(padx=10, pady=10, fill="x")

        # Results Table
        self.tree = ttk.Treeview(root, columns=("Combination", "Total Reach (%)", "Concept", "Total Reach of Concept (%)", "Incremental Reach (%)"), show='headings')
        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.pack(pady=10, fill="both", expand=True)
        self.tree.bind("<ButtonRelease-1>", self.on_combination_select)

        # Chart Frame
        self.chart_frame = tk.Frame(root)
        self.chart_frame.pack(pady=10, fill="both", expand=True)

        self.combinations_data = {}

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.data = pd.read_excel(file_path).fillna(False)
                self.show_filters()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    def show_filters(self):
        for widget in self.filters_frame.winfo_children():
            widget.destroy()

        self.columns = self.data.columns.tolist()
        self.check_vars = {}
        
        row_frame = tk.Frame(self.filters_frame)
        row_frame.pack(fill="x")

        for col in self.columns:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(row_frame, text=col, variable=var)
            chk.pack(side="left", padx=5)
            self.check_vars[col] = var

    def clear_chart(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

    def plot_waterfall_chart(self, combination, incremental_reach):
        self.clear_chart()
        
        concepts = list(combination)
        fig, ax = plt.subplots(figsize=(12, 4))  # Reduced height

        values = [sum(incremental_reach.values())] + [incremental_reach[c] for c in concepts]
        labels = [f"Total: {', '.join(combination)}"] + concepts
        colors = ['blue'] + ['green'] * len(concepts)

        bars = ax.bar(labels, values, color=colors)
        for i, bar in enumerate(bars):
            height = bar.get_height()
            offset = 5 if height < 90 else -10  # Adjust label position dynamically
            ax.text(bar.get_x() + bar.get_width()/2, height + offset, f"{height:.2f}%", ha='center', va='bottom', fontsize=10, color='black')

        ax.set_ylim(0, max(values) + 15)  # Extend Y-axis limit to avoid overlap
        ax.set_title(f"Waterfall Chart for {', '.join(combination)}")
        ax.set_ylabel("Reach (%)")
        plt.xticks(rotation=45)
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def run_turf_analysis(self):
        selected_columns = [col for col, var in self.check_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "Please select at least one column for analysis.")
            return

        subset_size = self.subset_size_entry.get()
        try:
            subset_size = int(subset_size)
        except ValueError:
            messagebox.showerror("Invalid Input", "Subset size must be an integer.")
            return

        if subset_size < 1 or subset_size > len(selected_columns):
            messagebox.showerror("Invalid Subset Size", "Subset size must be between 1 and the number of selected columns.")
            return

        # Clear previous results
        for row in self.tree.get_children():
            self.tree.delete(row)

        self.combinations_data.clear()
        selected_data = self.data[selected_columns].astype(bool)

        for combination in comb(selected_columns, subset_size):
            combined_data = selected_data[list(combination)].any(axis=1)
            total_reach_combination = combined_data.mean() * 100

            incremental_reach = {}
            previous_reach = 0

            for i, col in enumerate(combination):
                sub_combination = combination[:i + 1]
                current_reach = selected_data[list(sub_combination)].any(axis=1).mean() * 100
                incremental_reach[col] = current_reach - previous_reach
                previous_reach = current_reach

                self.tree.insert("", "end", values=(
                    ", ".join(combination),
                    f"{total_reach_combination:.2f}%",
                    col,
                    f"{selected_data[col].mean() * 100:.2f}%",
                    f"{incremental_reach[col]:.2f}%"
                ))
            
            self.combinations_data[", ".join(combination)] = incremental_reach

    def on_combination_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item[0], 'values')
            combination_key = item_values[0]
            if combination_key in self.combinations_data:
                self.plot_waterfall_chart(combination_key.split(", "), self.combinations_data[combination_key])

if __name__ == "__main__":
    root = tk.Tk()
    app = TurfAnalysisApp(root)
    root.mainloop()