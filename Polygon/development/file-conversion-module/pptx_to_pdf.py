import os
import tkinter as tk
from tkinter import filedialog, messagebox
from aspose.slides import Presentation
from aspose.slides.export import SaveFormat

def pptx_to_pdf():
    # Hide the root window
    root = tk.Tk()
    root.withdraw()

    # Select the PPTX file
    pptx_file = filedialog.askopenfilename(
        title="Select a PPTX file",
        filetypes=[("PPTX files", "*.pptx")]
    )
    if not pptx_file:
        messagebox.showinfo("No File Selected", "Please select a PPTX file.")
        return

    # Set output path to Downloads folder
    output_path = os.path.join(
        os.path.expanduser("~"),
        "Downloads",
        os.path.splitext(os.path.basename(pptx_file))[0] + ".pdf"
    )

    try:
        # Load presentation and save as PDF
        presentation = Presentation(pptx_file)
        presentation.save(output_path, SaveFormat.PDF)
        
        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    pptx_to_pdf()
