import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
import threading
import openpyxl
import pygetwindow as gw
import time


class ApplicationUsageTracker:
    def __init__(self, root):
        self.root = root
        self.root.title("Screen Time Analysis")
        self.root.geometry("400x600")
        self.root.configure(bg='#001B38')

        self.running = False
        self.start_time = None
        self.total_usage = {}
        self.elapsed_time = timedelta(seconds=0)

        self.start_button = ttk.Button(self.root, text="Start Analysis",
                                       command=self.start_analysis)
        self.stop_button = ttk.Button(self.root, text="Stop Analysis", command=self.stop_analysis)
        self.export_button = ttk.Button(self.root, text="Export to Excel",
                                        command=self.export_to_excel)

        self.start_button.configure(style='TButton.TButton')
        self.stop_button.configure(style='TButton.TButton')
        self.export_button.configure(style='TButton.TButton')

        self.time_label = ttk.Label(self.root, text="Time Elapsed: 0:00:00",
                                    background='#001B38', foreground='#8DB7C9')
        self.time_label.config(font=("Times New Roman", 11))

        self.result_label = tk.Label(self.root, text="", anchor="w", justify="left",
                                     bg='#001B38', foreground='#8DB7C9')
        self.result_label.config(font=("Times New Roman", 12))

        self.start_button.pack(pady=10)
        self.stop_button.pack(pady=10)
        self.export_button.pack(pady=10)
        self.time_label.pack(pady=10)
        self.result_label.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    def start_analysis(self):
        if not self.running:
            self.start_time = datetime.now() - self.elapsed_time
            self.running = True
            self.thread = threading.Thread(target=self.update_usage)
            self.thread.start()

    def stop_analysis(self):
        if self.running:
            self.running = False
            self.thread.join()
            self.elapsed_time += datetime.now() - self.start_time
            messagebox.showinfo("Analysis Finished", "Screen time analysis has been stopped.")

    def update_usage(self):
        while self.running:
            current_app = self.get_active_application()
            if current_app:
                self.total_usage[current_app] = self.total_usage.get(current_app, 0) + 1
            elapsed_time = datetime.now() - self.start_time + self.elapsed_time
            hours, remainder = divmod(int(elapsed_time.total_seconds()), 3600)
            minutes, seconds = divmod(remainder, 60)
            time_string = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            self.time_label.config(text=f"Time Elapsed: {time_string}")

            result_text = "\n".join([f"{app}: {self.format_time(seconds)}" for app, seconds in
                                     self.total_usage.items()])
            self.result_label.config(text=result_text)

            self.root.update()
            time.sleep(1)

    def get_active_application(self):
        try:
            active_window = gw.getActiveWindow()
            if active_window and active_window.title != "Application Usage Tracker":
                return active_window.title
        except Exception as e:
            pass

    def format_time(self, seconds):
        hours, remainder = divmod(seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def export_to_excel(self):
        if self.running:
            messagebox.showerror("ERROR!", "Stop the analysis before exporting to Excel.")
        elif self.total_usage:
            options = {
                "defaultextension": ".xlsx",
                "filetypes": [("Excel files", "*.xlsx"), ("All files", "*.*")],
                "initialfile": "screen_time_report.xlsx"
            }
            file_path = filedialog.asksaveasfilename(**options)

            if file_path:
                workbook = openpyxl.Workbook()
                worksheet = workbook.active
                worksheet.title = "Screen Time Analysis Report"
                worksheet["A1"] = "Applications"
                worksheet["B1"] = "Time Spent"
                row = 2
                for app, usage_seconds in self.total_usage.items():
                    time_string = self.format_time(usage_seconds)
                    worksheet.cell(row=row, column=1, value=app)
                    worksheet.cell(row=row, column=2, value=time_string)
                    row += 1
                workbook.save(file_path)
        else:
            messagebox.showerror("ERROR", "No data to export. Please start and stop the analysis "
                                          "first.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ApplicationUsageTracker(root)

    style = ttk.Style()
    style.configure('TButton.TButton', foreground='black', background='black', padding=(8, 8),
                    font=("Poppins", 11))

    root.mainloop()