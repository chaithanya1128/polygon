import os
import tkinter as tk
from tkinter import filedialog, messagebox
from aspose.slides import Presentation

def pdf_to_pptx():
    # Hide the root window
    root = tk.Tk()
    root.withdraw()

    # Select the PDF file
    pdf_file = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_file:
        messagebox.showinfo("No File Selected", "Please select a PDF file.")
        return

    # Set output path to Downloads folder
    output_path = os.path.join(os.path.expanduser("~"), "Downloads", os.path.splitext(os.path.basename(pdf_file))[0] + ".pptx")

    try:
        # Load PDF and save as PPTX
        presentation = Presentation()
        presentation.slides.add_from_pdf(pdf_file)
        presentation.save(output_path)
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    pdf_to_pptx()
