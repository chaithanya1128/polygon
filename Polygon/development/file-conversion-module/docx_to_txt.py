import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import os


def convert_docx_to_txt(file_paths, output_dir):
    """
    Convert selected DOCX files to TXT files and save them in the selected directory.
    """
    try:
        converted_count = 0
        for file_path in file_paths:
            if not file_path.lower().endswith(".docx"):
                continue  # Skip non-DOCX files

            # Set the output file path
            output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + ".txt")

            # Perform conversion
            doc = Document(file_path)
            with open(output_path, 'w', encoding='utf-8') as txt_file:
                for paragraph in doc.paragraphs:
                    txt_file.write(paragraph.text + '\n')

            converted_count += 1

        # Notify the user about the success
        messagebox.showinfo("Success", f"{converted_count} file(s) converted successfully! Saved at: {output_dir}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


def open_file_dialog():
    """
    Open a file dialog to select DOCX files and call the convert function.
    """
    # Open a file dialog to select DOCX files
    file_paths = filedialog.askopenfilenames(
        title="Select DOCX file(s)",
        filetypes=[("DOCX files", "*.docx")]
    )

    if not file_paths:
        messagebox.showinfo("No File Selected", "Please select at least one DOCX file to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Call the conversion function
    convert_docx_to_txt(file_paths, output_dir)


def create_gui():
    """
    Create the main GUI interface for the DOCX to TXT converter application.
    """
    # Set up the main application window
    root = tk.Tk()
    root.title("DOCX to TXT Converter")
    root.geometry("500x300")
    root.resizable(False, False)

    # Style the application
    style = ttk.Style(root)
    style.theme_use("clam")  # Use a clean theme
    style.configure("TLabel", font=("Arial", 12))
    style.configure("TButton", font=("Arial", 12))
    style.configure("Header.TLabel", font=("Arial", 16, "bold"), anchor="center")

    # Header Label
    header_label = ttk.Label(root, text="DOCX to TXT Converter", style="Header.TLabel")
    header_label.pack(pady=30)

    # Button to Open Files
    select_button = ttk.Button(
        root, text="Select DOCX File(s)", command=open_file_dialog, style="TButton"
    )
    select_button.pack(pady=20)

    # Footer Text
    footer_label = ttk.Label(
        root, text="Choose DOCX files and convert them into TXT files easily.", anchor="center"
    )
    footer_label.pack(pady=30)

    # Run the main application loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()
