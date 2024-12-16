import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import os

def convert_docx_to_txt():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()
    
    # Open a file dialog to select a DOCX file
    file_path = filedialog.askopenfilename(
        title="Select a DOCX file",
        filetypes=[("DOCX files", "*.docx")]
    )
    
    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a DOCX file to proceed.")
        return
    
    # Set the output file path
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(file_path))[0] + ".txt")
    
    try:
        # Perform conversion
        doc = Document(file_path)
        with open(output_path, 'w', encoding='utf-8') as txt_file:
            for paragraph in doc.paragraphs:
                txt_file.write(paragraph.text + '\n')
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    convert_docx_to_txt()
