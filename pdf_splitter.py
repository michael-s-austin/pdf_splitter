import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Extractor")
        # Set a fixed window size
        self.root.geometry("550x250")
        self.root.resizable(False, False)

        # 1. Input field and Browse Button
        tk.Label(root, text="Source PDF:").grid(row=0, column=0, padx=10, pady=15, sticky="e")
        self.file_entry = tk.Entry(root, width=45)
        self.file_entry.grid(row=0, column=1, padx=5, pady=15)
        
        self.browse_btn = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=10, pady=15)

        # 2. Page Range Inputs
        tk.Label(root, text="Start Page:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.start_entry = tk.Entry(root, width=10)
        self.start_entry.grid(row=1, column=1, sticky="w", padx=5)

        tk.Label(root, text="Stop Page:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.end_entry = tk.Entry(root, width=10)
        self.end_entry.grid(row=2, column=1, sticky="w", padx=5)

        # 3. New File Name Input
        tk.Label(root, text="New File Name:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.out_entry = tk.Entry(root, width=30)
        self.out_entry.grid(row=3, column=1, sticky="w", padx=5)
        self.out_entry.insert(0, "extracted_pages.pdf") # Default value

        # 4 & 5. Go and Quit Buttons
        self.go_btn = tk.Button(root, text="Go", width=10, command=self.execute_extraction, bg="#4CAF50", fg="white")
        self.go_btn.grid(row=4, column=1, sticky="w", padx=5, pady=20)

        self.quit_btn = tk.Button(root, text="Quit", width=10, command=self.root.destroy)
        self.quit_btn.grid(row=4, column=2, padx=10, pady=20)

    def browse_file(self):
        """Opens a file dialog to select a PDF and populates the entry field."""
        filepath = filedialog.askopenfilename(
            title="Select a PDF",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filepath)

    def execute_extraction(self):
        """Validates inputs and processes the PDF extraction."""
        input_path = self.file_entry.get().strip()
        start_str = self.start_entry.get().strip()
        end_str = self.end_entry.get().strip()
        out_name = self.out_entry.get().strip()

        # Validation: File exists
        if not os.path.exists(input_path):
            messagebox.showerror("Error", "The specified source file does not exist.")
            return

        # Validation: Valid integers for pages
        try:
            start_page = int(start_str)
            end_page = int(end_str)
        except ValueError:
            messagebox.showerror("Error", "Page range must be valid numbers.")
            return

        # Load document to check page bounds
        try:
            reader = PdfReader(input_path)
            total_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read PDF. Ensure it is not encrypted.\nDetails: {e}")
            return

        # Validation: Page ranges
        if start_page < 1 or end_page > total_pages or start_page > end_page:
            messagebox.showerror("Error", f"Invalid page range.\nDocument has {total_pages} pages.")
            return

        # Handle output filename extension
        if not out_name.lower().endswith('.pdf'):
            out_name += '.pdf'

        # Set output path to the same directory as the source file
        output_dir = os.path.dirname(os.path.abspath(input_path))
        output_path = os.path.join(output_dir, out_name)

        # Process extraction
        writer = PdfWriter()
        try:
            for i in range(start_page - 1, end_page):
                writer.add_page(reader.pages[i])

            with open(output_path, "wb") as out_file:
                writer.write(out_file)
                
            messagebox.showinfo("Success", f"Extracted pages {start_page} to {end_page}.\nSaved as:\n{output_name}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{e}")

# Application Entry Point
if __name__ == "__main__":
    # Create the main window and start the event loop
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()