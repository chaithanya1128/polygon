import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from comtypes.client import CreateObject


def convert_pptx_to_pdf(file_paths, output_dir):
    """
    Convert selected PPTX files to PDF and save them in the specified output directory.
    """
    try:
        # Initialize PowerPoint COM object
        powerpoint = CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1
        converted_count = 0

        for file_path in file_paths:
            if not file_path.lower().endswith(".pptx"):
                continue  # Skip non-PPTX files

            try:
                # Set the output file path
                pdf_output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + ".pdf")

                # Open the presentation and save it as PDF
                presentation = powerpoint.Presentations.Open(file_path, WithWindow=False)
                presentation.SaveAs(pdf_output_path, 32)  # 32 is the format for PDFs
                presentation.Close()
                converted_count += 1
            except Exception as file_error:
                messagebox.showerror("File Conversion Error", f"Error converting {file_path}: {file_error}")

        # Quit PowerPoint
        powerpoint.Quit()

        if converted_count > 0:
            messagebox.showinfo("Success", f"{converted_count} file(s) converted successfully! Saved at: {output_dir}")
        else:
            messagebox.showwarning("No Files Converted", "No files were successfully converted.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


def open_file_dialog():
    """
    Open a file dialog to select PPTX files and call the convert function.
    """
    # Open a file dialog to select PPTX files
    file_paths = filedialog.askopenfilenames(
        title="Select PowerPoint File(s)",
        filetypes=[("PowerPoint files", "*.pptx")]
    )

    if not file_paths:
        messagebox.showinfo("No File Selected", "Please select at least one PPTX file to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Call the conversion function
    convert_pptx_to_pdf(file_paths, output_dir)


def create_gui():
    """
    Create the main GUI interface for the PPTX to PDF converter application.
    """
    # Set up the main application window
    root = tk.Tk()
    root.title("PPTX to PDF Converter")
    root.geometry("500x300")
    root.resizable(False, False)

    # Style the application
    style = ttk.Style(root)
    style.theme_use("clam")  # Use a clean theme
    style.configure("TLabel", font=("Arial", 12))
    style.configure("TButton", font=("Arial", 12))
    style.configure("Header.TLabel", font=("Arial", 16, "bold"), anchor="center")

    # Header Label
    header_label = ttk.Label(root, text="PPTX to PDF Converter", style="Header.TLabel")
    header_label.pack(pady=30)

    # Button to Open Files
    select_button = ttk.Button(
        root, text="Select PowerPoint File(s)", command=open_file_dialog, style="TButton"
    )
    select_button.pack(pady=20)

    # Footer Text
    footer_label = ttk.Label(
        root, text="Choose PPTX files and convert them into PDF files easily.", anchor="center"
    )
    footer_label.pack(pady=30)

    # Run the main application loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()
