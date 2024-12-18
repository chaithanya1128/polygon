import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def convert_csv_to_xlsx():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select a CSV file
    file_path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv")]
    )

    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a CSV file to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Set the output XLSX file path
    output_file_path = filedialog.asksaveasfilename(
        title="Save XLSX As",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not output_file_path:
        messagebox.showinfo("No File Selected", "Please choose a location to save the XLSX file.")
        return

    try:
        # Load CSV file and convert it to an Excel file
        data = pd.read_csv(file_path)
        data.to_excel(output_file_path, index=False)

        messagebox.showinfo("Success", f"CSV file converted successfully to {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


if __name__ == "__main__":
    convert_csv_to_xlsx()
