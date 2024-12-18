import os
import tkinter as tk
from tkinter import filedialog, messagebox
import comtypes.client

def docx_to_pdf():
    # Hide the root window
    root = tk.Tk()
    root.withdraw()
    
    # Select the DOCX file
    docx_file = filedialog.askopenfilename(
        title="Select a DOCX file",
        filetypes=[("DOCX files", "*.docx")]
    )
    if not docx_file:
        messagebox.showinfo("No File Selected", "Please select a DOCX file.")
        return

    # Set output path to Downloads folder
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(docx_file))[0] + ".pdf")

    try:
        # Open Word application
        word = comtypes.client.CreateObject("Word.Application")
        doc = word.Documents.Open(docx_file)
        # Save as PDF
        doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        word.Quit()

        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    docx_to_pdf()
