import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def convert_xlsx_to_csv():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()
    
    # Open a file dialog to select an XLSX file
    file_path = filedialog.askopenfilename(
        title="Select an XLSX file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    
    if not file_path:
        messagebox.showinfo("No File Selected", "Please select an XLSX file to proceed.")
        return
    
    # Set the output file path
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(file_path))[0] + ".csv")
    
    try:
        # Perform conversion
        data = pd.read_excel(file_path)
        data.to_csv(output_path, index=False)
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    convert_xlsx_to_csv()
