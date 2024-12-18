import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def convert_csv_to_xlsx():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()
    
    # Open a file dialog to select a CSV file
    file_path = filedialog.askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV files", "*.csv")]
    )
    
    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a CSV file to proceed.")
        return
    
    # Set the output file path
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(file_path))[0] + ".xlsx")
    
    try:
        # Perform conversion
        data = pd.read_csv(file_path)
        data.to_excel(output_path, index=False)
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    convert_csv_to_xlsx()
