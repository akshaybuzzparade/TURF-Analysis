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
        self.root.geometry("1000x700")

        self.upload_button = tk.Button(root, text="Upload Excel File", command=self.load_file)
        self.upload_button.pack(pady=10)

        self.filters_frame = tk.LabelFrame(root, text="Filters", padx=10, pady=10)
        self.filters_frame.pack(padx=10, pady=10, fill="x")

        tk.Label(root, text="Subset Size:").pack()
        self.subset_size_entry = tk.Entry(root)
        self.subset_size_entry.pack(pady=5)

        self.run_analysis_button = tk.Button(root, text="Run TURF Analysis", command=self.run_turf_analysis)
        self.run_analysis_button.pack(pady=10)

        self.results_text = tk.Text(root, height=15, width=110)
        self.results_text.pack(pady=10)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            try:
                self.data = pd.read_excel(file_path)
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

    def plot_waterfall_chart(self, combination, total_reach, incremental_reach):
        concepts = list(combination)
        fig, ax = plt.subplots(figsize=(10, 6))

        values = [sum(incremental_reach.values())] + [incremental_reach[c] for c in concepts]
        labels = [f"Total: {', '.join(combination)}"] + concepts
        colors = ['blue'] + ['green'] * len(concepts)

        ax.bar(labels, values, color=colors)
        for i, value in enumerate(values):
            ax.text(i, value + 0.5, f"{value:.2f}%", ha='center', va='bottom')

        ax.set_title(f"Waterfall Chart for {', '.join(combination)}")
        ax.set_ylabel("Reach (%)")
        plt.xticks(rotation=45)
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.root)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=5)

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

        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"{'Combination':<25}{'Total Reach (%)':<20}{'Concept':<25}{'Total Reach of Concept (%)':<30}{'Incremental Reach (%)'}\n")
        self.results_text.insert(tk.END, "-" * 120 + "\n")

        selected_data = self.data[selected_columns]

        for combination in comb(selected_columns, subset_size):
            combined_data = selected_data[list(combination)].any(axis=1)
            total_reach_combination = combined_data.mean() * 100
            self.results_text.insert(tk.END, f"{', '.join(combination):<25}{total_reach_combination:<20.2f}%")

            incremental_reach = {}
            previous_reach = 0

            for i, col in enumerate(combination):
                sub_combination = combination[:i + 1]
                current_reach = selected_data[list(sub_combination)].any(axis=1).mean() * 100
                incremental_reach[col] = current_reach - previous_reach
                previous_reach = current_reach

                self.results_text.insert(tk.END, f"\n{'':<25}{'':<20}{col:<25}{selected_data[col].mean() * 100:<30.2f}{incremental_reach[col]:<20.2f}%")

            self.results_text.insert(tk.END, "\n")
            self.plot_waterfall_chart(combination, {col: selected_data[col].mean() * 100 for col in combination}, incremental_reach)

if __name__ == "__main__":
    root = tk.Tk()
    app = TurfAnalysisApp(root)
    root.mainloop()
