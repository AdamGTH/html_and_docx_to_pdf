from tkinter import *
from tkinter.filedialog import askopenfilenames
from tkinter.messagebox import showerror, showinfo
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
    path_to_libre_engine = ""
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        self.geometry("650x500")
        self.configure(bg="gray")
        self.title("Convert to Pdf from .docx and .html")
        
        self.frame_button = Frame(self)
        self.frame_button.pack(pady=5)
        self.frame_text = Frame(self, relief="groove")
        self.frame_text.pack(pady=5)
        
        self.btn_select = Button(self.frame_button,
        text="SELECT FILES", width=15, command=self.select_files
        )
        
        self.btn = Button(self.frame_button, text="MAKE PDF", width=15,
                command=self.thread_to_bar, state="disabled")

        self.btn_select.grid(padx=5, pady=5, column=0, row=0)
        self.btn.grid(padx=5 ,pady=5, column=1, row=0)

        self.info_lab = Label(self.frame_text, font=("Arial Black", "10"), fg="blue", text="No files")
        self.text = Text(self.frame_text, wrap='none', bg='black', fg="white")

        self.info_lab.grid(padx=5, pady=5, row=0, column=0)
        self.text.grid(padx=5, column=0, row=1)
        
    def select_files(self):
        self.paths_tuple = askopenfilenames(
            parent=self, filetypes=[
                ("Text Files", "*.htm"), ("Text Files", "*.html"), ("Text Files", "*.docx"), ("Text Files", "*.txt")]
        )
        
        self.path_this = os.getcwd()
        self.ilosc = len(self.paths_tuple)
        wantedIdxStart = 0
        for chars in self.paths_tuple:
            for i, zn in enumerate(chars):
                if zn == '/':
                    wantedIdxStart = i
        
        wantedIdxEnd = -4
        
        if self.paths_tuple[0].find("html") > 0 or self.paths_tuple[0].find("docx") > 0:
            wantedIdxEnd = -5
        
        self.list_names_files = [name[wantedIdxStart+1:wantedIdxEnd] for name in self.paths_tuple]
        self.btn.configure(state="normal")
        self.info_lab.configure(text=f"loaded {len(self.paths_tuple)} files")
        
        for val in self.list_names_files:
            self.text.insert(INSERT, val + "\n")
            
        # pobieranie danych z pliku config
        try:
            with open("config.txt", "r") as fr:
                for line in fr:
                    if "path_to_libre_engine" in line:
                        self.path_to_libre_engine = line.split("=")[1].strip()
        except FileNotFoundError:
            showinfo("MESSAGE", "NOT FOUND config.txt with path to swriter.exe, used default path: C:/Program Files/LibreOffice/program/swriter.exe")
            self.path_to_libre_engine = "C:/Program Files/LibreOffice/program/swriter.exe"
      
    def to_create(self):
        subprocess.call(["mkdir", "pdfs"])              
        if "doc" in self.paths_tuple[0] or "docx" in self.paths_tuple[0]:
            # os.system(f"\"C:/Program Files/LibreOffice/program/swriter.exe\" --headless --convert-to pdf --outdir {path_to_save} {doc_name_docx}")
            for idx, file in enumerate(self.paths_tuple):
                commandStrings = [self.path_to_libre_engine, "--headless", "--convert-to", "pdf", "--outdir", "pdfs/", file]
                try:
                    retCode = subprocess.call(commandStrings)
                    self.add_progress_to_text(idx)
                except:
                    showerror("WARNING !", f"{retCode}")
        
        else:
            # generowanie pdfow z protokolow .htm
            for idx, file in enumerate(self.paths_tuple):
                try:
                    pdfkit.from_file(
                file,
                f"pdfs/{self.list_names_files[idx]}.pdf",
                # verbose=True,
                options={"enable-local-file-access": True},
            )
                    self.add_progress_to_text(idx)
                except:
                    showerror("WARNING !", "Sometning is impossible to do !!!")
                    return   
                    
        # generowanie zbiorczego pliku pdf
        pdfs_reader = []
        for file in self.list_names_files:
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