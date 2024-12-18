import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image


def convert_image_to_pdf():
    # Initialize tkinter root
    root = tk.Tk()
    root.withdraw()

    # Open a file dialog to select image files
    file_paths = filedialog.askopenfilenames(
        title="Select Image File(s)",
        filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp")]
    )

    if not file_paths:
        messagebox.showinfo("No File Selected", "Please select at least one image to proceed.")
        return

    # Ask the user to choose the output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory")

    if not output_dir:
        messagebox.showinfo("No Directory Selected", "Please select an output directory to proceed.")
        return

    # Set the output PDF file path
    pdf_output_path = filedialog.asksaveasfilename(
        title="Save PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not pdf_output_path:
        messagebox.showinfo("No File Selected", "Please choose a location to save the PDF.")
        return

    try:
        # Convert the images to PDF
        image_list = [Image.open(file_path) for file_path in file_paths]
        image_list[0].save(pdf_output_path, save_all=True, append_images=image_list[1:], resolution=100.0, quality=95,
                           optimize=True)

        messagebox.showinfo("Success", f"Image(s) converted successfully to {pdf_output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during conversion: {e}")


if __name__ == "__main__":
    convert_image_to_pdf()
