import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter import simpledialog
from PIL import Image
from fpdf import FPDF
import os
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_path
import pandas as pd

class FileConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")
        self.create_gui()

    def create_gui(self):
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, mode="determinate")
        self.progress_bar.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="we")

        self.conversion_type_label = ttk.Label(self.root, text="Select Conversion Type:")
        self.conversion_type_label.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")

        self.conversion_type = tk.StringVar()
        self.conversion_type.set("Text to PDF")
        self.conversion_type_menu = ttk.Combobox(self.root, textvariable=self.conversion_type,
                                                 values=["Text to PDF", "Image to PDF", "PDF to Image", "Merge PDFs", "Lock PDF with Password"])
        self.conversion_type_menu.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="we")
        self.conversion_type_menu.bind("<<ComboboxSelected>>", self.on_conversion_type_selected)

        self.select_files_btn = ttk.Button(self.root, text="Select Files", command=self.select_files)
        self.select_files_btn.grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.comment_box = ttk.Label(self.root, text="", wraplength=400)
        self.comment_box.grid(row=3, column=1, padx=10, pady=5, sticky="e")

        self.convert_btn = ttk.Button(self.root, text="Convert", command=self.convert_files)
        self.convert_btn.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="we")

        self.files = []
        self.pdf = None

    def on_conversion_type_selected(self, event):
        # Clear selected files, PDF, and update the progress bar
        self.files = []
        self.pdf = None
        self.progress_var.set(0)
        self.comment_box.config(text="")

    def select_files(self):
        self.files = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=[
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("Image files", "*.jpg;*.jpeg;*.png"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        self.comment_box.config(text="\n".join(self.files))

    def convert_files(self):
        conversion_type = self.conversion_type.get()
        if not self.files:
            messagebox.showerror("Error", "No files selected for conversion.")
            return

        if conversion_type == "Text to PDF":
            self.convert_text_to_pdf()
        elif conversion_type == "Image to PDF":
            self.convert_image_to_pdf()
        elif conversion_type == "PDF to Image":
            self.convert_pdf_to_image()
        elif conversion_type == "Merge PDFs":
            self.merge_pdfs()
        elif conversion_type == "Lock PDF with Password":
            self.lock_pdf_with_password()

    def convert_text_to_pdf(self):
        pdf_name = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if pdf_name:
            try:
                self.pdf = FPDF()
                self.pdf.add_page()
                self.pdf.set_font("Arial", size=12)
                total_files = len(self.files)

                for i, file in enumerate(self.files):
                    with open(file, "r") as text_file:
                        for line in text_file:
                            self.pdf.cell(0, 10, txt=line, ln=True)

                    # Update progress bar
                    progress = (i + 1) / total_files * 100
                    self.progress_var.set(int(progress))
                    self.root.update_idletasks()

                self.pdf.output(pdf_name)
                messagebox.showinfo("Success", "Text files have been successfully converted to PDF.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert text to PDF.\nError: {str(e)}")

    def convert_image_to_pdf(self):
        pdf_name = filedialog.asksaveasfilename(
            title="Save PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if pdf_name:
            try:
                total_images = len(self.files)
                if total_images > 0:
                    img = Image.open(self.files[0])
                    img.save(pdf_name, "PDF", resolution=100.0, save_all=True, append_images=self.files[1:])
                    messagebox.showinfo("Success", "Images have been successfully converted to PDF.")
                else:
                    messagebox.showerror("Error", "No images selected for conversion.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert images to PDF.\nError: {str(e)}")

    def convert_pdf_to_image(self):
        output_folder = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if output_folder:
            try:
                total_pdfs = len(self.files)
                if total_pdfs > 0:
                    for i, pdf_file in enumerate(self.files):
                        pages = convert_from_path(pdf_file, dpi=300)
                        pdf_filename = os.path.basename(pdf_file)
                        pdf_filename_noext = os.path.splitext(pdf_filename)[0]

                        for page_num, image in enumerate(pages):
                            image.save(os.path.join(output_folder, f"{pdf_filename_noext}_page_{page_num + 1}.png"),
                                       "PNG")

                        # Update progress bar
                        progress = (i + 1) / total_pdfs * 100
                        self.progress_var.set(int(progress))
                        self.root.update_idletasks()

                    messagebox.showinfo("Success", "PDFs have been successfully converted to images.")
                else:
                    messagebox.showerror("Error", "No PDFs selected for conversion.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to convert PDFs to images.\nError: {str(e)}")

    def merge_pdfs(self):
        pdf_name = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if pdf_name:
            try:
                pdf_writer = PdfWriter()
                for pdf_file in self.files:
                    pdf_reader = PdfReader(pdf_file)
                    for page_num in range(pdf_reader.getNumPages()):
                        pdf_writer.addPage(pdf_reader.pages(page_num))
                with open(pdf_name, "wb") as output_pdf:
                    pdf_writer.write(output_pdf)
                messagebox.showinfo("Success", "PDFs have been successfully merged.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs.\nError: {str(e)}")

    def lock_pdf_with_password(self):
        pdf_name = filedialog.asksaveasfilename(
            title="Save Locked PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if pdf_name:
            password = simpledialog.askstring("Password", "Enter the password to lock the PDF:")
            if password:
                try:
                    reader = PdfReader(self.files[0])
                    writer = PdfWriter()
                    writer.append_pages_from_reader(reader)
                    writer.encrypt(password)

                    with open(pdf_name, "wb") as out_file:
                        writer.write(out_file)
                    messagebox.showinfo("Success", "PDF locked with password and saved to: " + pdf_name)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to lock PDF with password.\nError: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverter(root)
    root.mainloop()
