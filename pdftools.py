# -*- coding: utf-8 -*-

from customtkinter import *
from tkinter import *
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile
from tkinter import messagebox
import os
from pdf2docx import Converter
from subprocess import run
import PyPDF2
import sys



# https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



root = CTk()
root.geometry("700x500")
root.title("PDF Tools App")
root.resizable(False, False)

def open_doc_window():
    app = CTk()
    app.geometry("350x200")
    app.title("Convert To Docs")
    app.columnconfigure(0, weight=1)

    title = CTkLabel(app, text="Select Pdf to Convert", font=('Lato', 16, 'bold'))
    title.grid(row=0, column=0, padx=20, pady=(20,10), columnspan=3)

    def open_file():
        file = askopenfile(parent=app, mode="rb", title="Choose a file", filetypes=[("Pdf file", "*.pdf")])
        if file:
            convert_to_word(file.name)

    def convert_to_word(pdf_path):
        output_path = "output.docx"
        try:
            with open(pdf_path, 'rb') as pdf_file, open(output_path, 'wb') as docx_file:
                cv = Converter(pdf_file)
                cv.convert(docx_file, start=0, end=None)
                cv.close()
            messagebox.showinfo("Conversion Complete", "PDF converted to Word successfully!")
            display_open_button(app)
        except Exception as e:
            messagebox.showerror("Conversion Error", f"An error occurred: {str(e)}")

    def display_open_button(app):
        open_doc_btn = CTkButton(app, text="Open Document", font=('Lato', 14, 'bold'), command=open_word_doc, fg_color="#149072", height=35)
        open_doc_btn.grid(row=2, column=0, padx=20, pady=10, columnspan=3)

    def open_word_doc():
        os.system('start output.docx')

    search = CTkButton(app, text="Select PDF", font=('Lato', 14, 'bold'), command=open_file, height=35)
    search.grid(row=1, column=0, columnspan=3, padx=20, pady=10)

    app.mainloop()

def open_text_window():
    app = CTk()
    app.geometry("350x200")
    app.title("Convert To Docs")
    app.columnconfigure(0, weight=1)

    title = CTkLabel(app, text="Select Pdf to Convert", font=('Lato', 16, 'bold'))
    title.grid(row=0, column=0, padx=20, pady=(20,10), columnspan=3)

    def open_file():
        def convert_to_text(pdf_path):
            output_path = "output_txt.txt"
            try:
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)
                    with open(output_path, 'w', encoding='utf-8') as txt_file:
                        for page_num in range(len(pdf_reader.pages)):
                            page = pdf_reader.pages[page_num]
                            text = page.extract_text()
                            txt_file.write(text)
                messagebox.showinfo("Conversion Complete", "PDF converted to text successfully!")
                display_open_button(app)
            except Exception as e:
                messagebox.showerror("Conversion Error", f"An error occurred: {str(e)}")

        file = askopenfile(parent=app, mode="rb", title="Choose a file", filetypes=[("Pdf file", "*.pdf")])
        if file:
            convert_to_text(file.name)

    def display_open_button(app):
        open_doc_txt = CTkButton(app, text="Open Text File", font=('Lato', 14, 'bold'), command=open_txt_file, fg_color="#149072", height=35)
        open_doc_txt.grid(row=2, column=0, padx=20, pady=10, columnspan=3)

    def open_txt_file():
        os.system('start output_txt.txt')

    search = CTkButton(app, text="Select PDF", font=('Lato', 14, 'bold'), command=open_file, height=35)
    search.grid(row=1, column=0, columnspan=3, padx=20, pady=10)

    app.mainloop()

def open_encrypt_window():
    app = CTk()
    app.geometry("350x250")
    app.title("Encrypt PDF")
    app.columnconfigure(0, weight=1)

    password = CTkEntry(app, placeholder_text="Enter password", show="*", font=('Lato', 14, 'bold'))
    password.grid(row=1, column=0, columnspan=3, padx=20, pady=(0,20))

    def encrypt_pdf(input_path, output_path):
        try:
            with open(input_path, 'rb') as input_file:
                pdf_reader = PyPDF2.PdfReader(input_file)
                pdf_writer = PyPDF2.PdfWriter()
                for page_num in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page_num])
                pdf_writer.encrypt(password.get())
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
            messagebox.showinfo("Encryption Complete", "PDF successfully Encrypted!")
            display_open_button(app)
        except Exception as e:
            messagebox.showerror("Encryption Error", f"An error occurred: {str(e)}")

    def open_file():
        file = askopenfile(parent=app, mode="rb", title="Choose a file", filetypes=[("Pdf file", "*.pdf")])
        if file:
            encrypt_pdf(file.name, "encrypted.pdf")

    def display_open_button(app):
        open_encrypt_pdf = CTkButton(app, text="Open Encrypted PDF", font=('Lato', 14, 'bold'), command=open_encrypt, fg_color="#149072", height=35)
        open_encrypt_pdf.grid(row=3, column=0, padx=20, pady=(20,20), columnspan=3)

    def open_encrypt():
        os.system('start encrypted.pdf')

    title = CTkLabel(app, text="Select PDF to Encrypt", font=('Lato', 16, 'bold'))
    title.grid(row=0, column=0, columnspan=3, padx=20, pady=20)

    search = CTkButton(app, text="Select PDF", font=('Lato', 14, 'bold'), command=open_file, height=35)
    search.grid(row=2, column=0, columnspan=3, padx=20, pady=(0,10))

    app.mainloop()

def open_merge_window():
    app = CTk()
    app.geometry("350x300")
    app.title("Merge PDFs")
    app.columnconfigure(0, weight=1)
    app.columnconfigure(1, weight=1)
    app.columnconfigure(2, weight=1)

    pdf1_path = None
    pdf2_path = None

    def merge_pdfs(pdf1_path, pdf2_path, output_path):
        with open(pdf1_path, 'rb') as pdf1_file, open(pdf2_path, 'rb') as pdf2_file:
            pdf1_reader = PyPDF2.PdfReader(pdf1_file)
            pdf2_reader = PyPDF2.PdfReader(pdf2_file)
            pdf_writer = PyPDF2.PdfWriter()
            for page_num in range(len(pdf1_reader.pages)):
                page = pdf1_reader.pages[page_num]
                pdf_writer.add_page(page)
            for page_num in range(len(pdf2_reader.pages)):
                page = pdf2_reader.pages[page_num]
                pdf_writer.add_page(page)
            with open(output_path, 'wb') as merged_pdf_file:
                pdf_writer.write(merged_pdf_file)
        messagebox.showinfo("Merge Complete", "PDFs merged successfully!")
        display_open_button(app)

    def select_pdf1():
        nonlocal pdf1_path
        file = askopenfile(parent=app, mode="rb", title="Choose the first PDF", filetypes=[("PDF files", "*.pdf")])
        if file:
            pdf1_path = file.name

    def select_pdf2():
        nonlocal pdf2_path
        file = askopenfile(parent=app, mode="rb", title="Choose the second PDF", filetypes=[("PDF files", "*.pdf")])
        if file:
            pdf2_path = file.name

    def merge_pdfs_wrapper():
        if pdf1_path and pdf2_path:
            output_path = "merged.pdf"
            merge_pdfs(pdf1_path, pdf2_path, output_path)
        else:
            messagebox.showerror("Error", "Please select both PDF files.")

    def display_open_button(app):
        open_merge_btn = CTkButton(app, text="Open Merged PDF", font=('Lato', 14, 'bold'), command=open_merge_pdf, height=35, width=200, fg_color="#149072")
        open_merge_btn.grid(row=3, column=0, padx=20, pady=20, columnspan=3)

    def open_merge_pdf():
        os.system('start merged.pdf')

    select_pdf1_btn = CTkButton(app, text="Select 1st PDF", font=('Lato', 14, 'bold'), command=select_pdf1, height=35)
    select_pdf1_btn.grid(row=1, column=0, padx=20, pady=10)

    select_pdf2_btn = CTkButton(app, text="Select 2nd PDF", font=('Lato', 14, 'bold'), command=select_pdf2, height=35)
    select_pdf2_btn.grid(row=1, column=1, padx=20, pady=10)

    merge_btn = CTkButton(app, text="Merge PDFs", font=('Lato', 14, 'bold'), command=merge_pdfs_wrapper, height=35)
    merge_btn.grid(row=2, column=0, padx=20, pady=(20,10), columnspan=3)

    title = CTkLabel(app, text="Select PDFs to Merge", font=('Lato', 16, 'bold'))
    title.grid(row=0, column=0, columnspan=3, padx=20, pady=20)

    app.mainloop()

# Main Frame
background_image = Image.open(resource_path("assets\\pattern.png"))
background_photo = ImageTk.PhotoImage(background_image)
background_label = Label(root, image=background_photo)
background_label.place(relwidth=1, relheight=1)

frame = CTkFrame(root, height=200, width=520)
frame.grid(row=0, column=0, rowspan=3, columnspan=3, padx=60, pady=40, sticky="nsew")
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Name Frame
name_frame = CTkFrame(frame, fg_color="transparent")
name = CTkLabel(name_frame, text="PDF Tools", font=('Lato', 24, 'bold'))
name.grid(row=0, column=0, padx=(20, 20), pady=(20, 20))
name_frame.grid(row=0, column=0, padx=(20, 20), pady=(20, 0))

frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)

# Tools Frame
tool_frame_1 = CTkFrame(frame, fg_color="transparent")

# Word Frame
word_frame = CTkFrame(tool_frame_1, fg_color="#2B6592")
word_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

word_btn = CTkButton(word_frame, text="", fg_color="#2B6592", hover_color="#2B6592", command=open_doc_window)
word_btn.grid(row=0, column=0)

word_icon = Image.open(resource_path("assets\\word.png"))
word_icon = ImageTk.PhotoImage(word_icon)
word_icon_label = Label(word_frame, image=word_icon, bg="#2B6592")
word_icon_label.grid(row=0, column=0, padx=(40, 40), pady=(20, 0))

word = CTkLabel(word_frame, text="PDF To Word", font=('Lato', 16, 'bold'))
word.grid(row=1, column=0, padx=(20, 20), pady=(0, 10))

# Text Frame
text_frame = CTkFrame(tool_frame_1, fg_color="#246091")
text_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

text_btn = CTkButton(text_frame, text="", width=90, fg_color="#246091", hover_color="#246091", command=open_text_window)
text_btn.grid(row=0, column=0)

text_icon = Image.open(resource_path("assets\\text.png"))
text_icon = ImageTk.PhotoImage(text_icon)
text_icon_label = Label(text_frame, image=text_icon, bg="#246091")
text_icon_label.grid(row=0, column=0, padx=(10, 10), pady=(20, 0))

text = CTkLabel(text_frame, text="PDF To Text", font=('Lato', 16, 'bold'))
text.grid(row=1, column=0, padx=(20, 20), pady=(10, 8))

# Encrypt Frame
pass_frame = CTkFrame(tool_frame_1, fg_color="#185967")
pass_frame.grid(row=0, column=2, padx=20, pady=20, sticky="nsew")

pass_btn = CTkButton(pass_frame, text="", width=90, fg_color="#185967", hover_color="#185967", command=open_encrypt_window)
pass_btn.grid(row=0, column=0)

pass_icon = Image.open(resource_path("assets\\pass.png"))
pass_icon = ImageTk.PhotoImage(pass_icon)
pass_icon_label = Label(pass_frame, image=pass_icon, bg="#185967")
pass_icon_label.grid(row=0, column=0, padx=(10, 10), pady=(20, 0))

password = CTkLabel(pass_frame, text="Encrypt PDF", font=('Lato', 16, 'bold'))
password.grid(row=1, column=0, padx=(20, 20), pady=(10, 10))

# Merge Frame
merge_frame = CTkFrame(tool_frame_1, fg_color="#802C45")
merge_frame.grid(row=1, column=1, padx=20, pady=20, sticky="nsew")

merge_btn = CTkButton(merge_frame, text="", width=90, fg_color="#802C45", hover_color="#802C45", command=open_merge_window)
merge_btn.grid(row=0, column=0)

merge_icon = Image.open(resource_path("assets\\merge.png"))
merge_icon = ImageTk.PhotoImage(merge_icon)
merge_icon_label = Label(merge_frame, image=merge_icon, bg="#802C45")
merge_icon_label.grid(row=0, column=0, padx=(10, 10), pady=(20, 0))

merge = CTkLabel(merge_frame, text="PDF To Merge", font=('Lato', 16, 'bold'))
merge.grid(row=1, column=0, padx=(20, 20), pady=(10, 10))

tool_frame_1.grid(row=1, column=0)

root.mainloop()
