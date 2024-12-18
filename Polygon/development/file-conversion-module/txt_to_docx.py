import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import os

def convert_txt_to_docx():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()
    
    # Open a file dialog to select a TXT file
    file_path = filedialog.askopenfilename(
        title="Select a TXT file",
        filetypes=[("Text files", "*.txt")]
    )
    
    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a TXT file to proceed.")
        return
    
    # Set the output file path
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(file_path))[0] + ".docx")
    
    try:
        # Perform conversion
        doc = Document()
        with open(file_path, 'r', encoding='utf-8') as txt_file:
            for line in txt_file:
                doc.add_paragraph(line.strip())
        doc.save(output_path)
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    convert_txt_to_docx()
