import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os


def convert_pdf_to_docx():
    # Open a file dialog to select a PDF file
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not file_path:
        messagebox.showinfo("No File Selected", "Please select a PDF file to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Set the output file path
    output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + ".docx")

    try:
        # Perform conversion
        cv = Converter(file_path)
        cv.convert(output_path)
        cv.close()

        messagebox.showinfo("Success", f"File converted successfully! Saved at: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


# Create the main application window
root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("400x200")
root.resizable(False, False)

# Header Label
header_label = tk.Label(root, text="PDF to DOCX Converter", font=("Arial", 16, "bold"))
header_label.pack(pady=20)

# Convert Button
convert_button = tk.Button(root, text="Select PDF to Convert", font=("Arial", 12), command=convert_pdf_to_docx)
convert_button.pack(pady=20)

# Footer Label
footer_label = tk.Label(root, text="Choose a PDF file and convert it into a DOCX file.", font=("Arial", 10))
footer_label.pack(pady=20)

# Run the application
root.mainloop()
