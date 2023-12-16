from tkinter import *
from tkinter.filedialog import askopenfilenames
from tkinter.messagebox import showerror
from tkinter import ttk
import pdfkit
import os
import subprocess
from PyPDF2 import PdfWriter, PdfReader
import threading
import datetime
from tkinter.scrolledtext import ScrolledText
from docxtpl import DocxTemplate
from pdf2docx import Converter


class WindowClass(Tk):
    paths_tuple = ''
    path_this = ''
    ilosc = 0
    list_names_files = []
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.geometry("800x500")
        self.configure(bg="gray")
        self.title("Convert to Pdf from .docx and .html")
        
        self.frame_button = Frame(self)
        self.frame_button.pack(pady=5)
        self.frame_text = Frame(self, relief="groove")
        self.frame_text.pack(pady=5)
        
        self.btn_select = Button(self.frame_button,
        text="Select Files", width=15, bg="dark green", command=self.select_files
        )
        
        self.btn = Button(self.frame_button, text="Make Pdf", width=15,
                command=self.thread_to_bar, state="disabled")

        self.btn_select.grid(padx=5, pady=5, column=0, row=0)
        self.btn.grid(padx=5 ,pady=5, column=1, row=0)

        self.info_lab = Label(self.frame_text, font=("Arial Black", "10"), fg="blue", text="No files")
        self.text = Text(self.frame_text, wrap='none', bg='black', fg="white")

        self.info_lab.grid(padx=5, pady=5, row=0, column=0)
        self.text.grid(padx=5, column=0, row=1)
        
    def select_files(self):
        global paths_tuple, path_this, ilosc, list_names_files
        paths_tuple = askopenfilenames(
            parent=self, filetypes=[
                ("Text Files", "*.htm"), ("Text Files", "*.html"), ("Text Files", "*.docx"), ("Text Files", "*.txt")]
        )
        
        path_this = os.getcwd()
        ilosc = len(paths_tuple)
        wantedIdxStart = 0
        for chars in paths_tuple:
            for i, zn in enumerate(chars):
                if zn == '/':
                    wantedIdxStart = i
        
        wantedIdxEnd = -4
        
        if paths_tuple[0].find("html") > 0 or paths_tuple[0].find("docx") > 0:
            wantedIdxEnd = -5
        
        list_names_files = [name[wantedIdxStart+1:wantedIdxEnd] for name in paths_tuple]
        self.btn.configure(state="normal")
        self.info_lab.configure(text=f"loaded {len(paths_tuple)} files")
        
        for val in list_names_files:
            self.text.insert(INSERT, val + "\n")
      
    def to_create(self):
        # pobieranie danych z pliku config
        with open("config.txt", "r") as fr:
            for line in fr:
                if "path_to_libre_engine" in line:
                    path_to_libre_engine = line.split("=")[1].strip()
                    
        if "doc" in paths_tuple[0] or "docx" in paths_tuple[0]:
            # os.system(f"\"C:/Program Files/LibreOffice/program/swriter.exe\" --headless --convert-to pdf --outdir {path_to_save} {doc_name_docx}")
            for idx, file in enumerate(paths_tuple):
                commandStrings = [path_to_libre_engine, "--headless", "--convert-to", "pdf", "--outdir", "pdfs/", file]
                retCode = subprocess.call(commandStrings)
                self.add_progress_to_text(idx)
        
        else:
            # generowanie pdfow z protokolow .htm
            for idx, file in enumerate(paths_tuple):
                pdfkit.from_file(
            file,
            f"pdfs/{list_names_files[idx]}.pdf",
            # verbose=True,
            options={"enable-local-file-access": True},
        )
                self.add_progress_to_text(idx)   
                    
        # generowanie zbiorczego pliku pdf
        pdfs_reader = []
        for file in list_names_files:
            pdfs_reader.append(PdfReader(open(f"pdfs/{file}.pdf", "rb")))
        
        output_pdf = PdfWriter()
        
        for idx, file in enumerate(pdfs_reader):
            for page in file.pages:
                output_pdf.add_page(page)

        with open("pdfs/all_doc.pdf", "wb") as out_stream:
            output_pdf.write(out_stream)
            
        self.text.insert(END, "SUCCESSFUL CONVERSION !!!")
              
    def thread_to_bar(self):
        th1 = threading.Thread(target=self.to_create)
        th1.start()
    
    def add_progress_to_text(self, idx):
        self.text.insert(f"{idx+1}.0", "OK! ") 
        self.text.tag_add("start", f"{idx+1}.0", f"{idx+1}.3") 
        #configuring a tag called start 
        self.text.tag_config("start", background="green", 
                        foreground="white")


main_win = WindowClass()
main_win.mainloop()