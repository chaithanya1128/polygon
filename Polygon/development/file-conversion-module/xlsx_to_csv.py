import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os


def convert_xlsx_to_csv(file_paths, output_dir):
    """
    Convert selected XLSX files to CSV and save them in the specified output directory.
    """
    try:
        converted_count = 0
        for file_path in file_paths:
            if not file_path.lower().endswith(".xlsx"):
                continue  # Skip non-XLSX files

            # Set the output file path
            output_path = os.path.join(output_dir, os.path.splitext(os.path.basename(file_path))[0] + ".csv")

            # Perform conversion
            data = pd.read_excel(file_path)
            data.to_csv(output_path, index=False)
            converted_count += 1

        # Notify the user about the success
        messagebox.showinfo("Success", f"{converted_count} file(s) converted successfully! Saved at: {output_dir}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


def open_file_dialog():
    """
    Open a file dialog to select XLSX files and call the convert function.
    """
    # Open a file dialog to select XLSX files
    file_paths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not file_paths:
        messagebox.showinfo("No File Selected", "Please select at least one XLSX file to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Call the conversion function
    convert_xlsx_to_csv(file_paths, output_dir)


def create_gui():
    """
    Create the main GUI interface for the XLSX to CSV converter application.
    """
    # Set up the main application window
    root = tk.Tk()
    root.title("XLSX to CSV Converter")
    root.geometry("500x300")
    root.resizable(False, False)

    # Style the application
    style = ttk.Style(root)
    style.theme_use("clam")  # Use a clean theme
    style.configure("TLabel", font=("Arial", 12))
    style.configure("TButton", font=("Arial", 12))
    style.configure("Header.TLabel", font=("Arial", 16, "bold"), anchor="center")

    # Header Label
    header_label = ttk.Label(root, text="XLSX to CSV Converter", style="Header.TLabel")
    header_label.pack(pady=30)

    # Button to Open Files
    select_button = ttk.Button(
        root, text="Select Excel File(s)", command=open_file_dialog, style="TButton"
    )
    select_button.pack(pady=20)

    # Footer Text
    footer_label = ttk.Label(
        root, text="Choose Excel files and convert them into CSV files easily.", anchor="center"
    )
    footer_label.pack(pady=30)

    # Run the main application loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()
