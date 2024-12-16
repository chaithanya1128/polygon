import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from pdf2docx import Converter
import os

def convert_pdf_to_docx():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    
    # Open a file dialog to select a PDF file
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    
    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a PDF file to proceed.")
        return
    
    # Set the output file path
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(file_path))[0] + ".docx")
    
    try:
        # Perform conversion
        cv = Converter(file_path)
        cv.convert(output_path)
        cv.close()
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    convert_pdf_to_docx()
