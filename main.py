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

paths_tuple = ''
path_this = ''
ilosc = 0
list_names_files = []
current_time = datetime.datetime.now()
oum_mans = []
with open("urzednicy.txt", "r", encoding="utf8") as fu:
    for man in fu:
        oum_mans.append(man)

def thread_to_bar():
    th1 = threading.Thread(target=to_create)
    th1.start()

# def progress_bar(actual_iterable):
#     global ilosc
#     pcnt = actual_iterable * 100 / ilosc
#     bar["value"] = pcnt
#     bar_val["text"] = f'{round(bar["value"],2)} %'
#     if bar["value"] == 100.0:
#        bar_val["text"] = "Finish !!!"

def select_files():
    global paths_tuple, path_this, ilosc, list_names_files
    paths_tuple = askopenfilenames(
        parent=root, filetypes=[
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
    btn.configure(state="normal")
    info_lab.configure(text=f"loaded {len(paths_tuple)} files")
    
    for val in list_names_files:
        text.insert(INSERT, val + "\n")
    
   
def to_create():
    path_to_save = ''
    path_to_confirm = ''
    path_to_template = ""
    idx_val = 0
    
      # pobieranie danych z pliku config
    with open("config.txt", "r") as fr:
        for line in fr:
            if "path_to_save" in line:
                path_to_save = line.split("=")[1].strip()
            if "path_to_libre_engine" in line:
                path_to_libre_engine = line.split("=")[1].strip()
                
    if "doc" in paths_tuple[0] or "docx" in paths_tuple[0]:
        # os.system(f"\"C:/Program Files/LibreOffice/program/swriter.exe\" --headless --convert-to pdf --outdir {path_to_save} {doc_name_docx}")
        for idx, file in enumerate(paths_tuple):
            commandStrings = [path_to_libre_engine, "--headless", "--convert-to", "pdf", "--outdir", "pdfs/", file]
            retCode = subprocess.call(commandStrings)
            # add tag using indices for the 
            # part of text to be highlighted
            text.insert(f"{idx+1}.0", "OK! ") 
            text.tag_add("start", f"{idx+1}.0", f"{idx+1}.3") 
            #configuring a tag called start 
            text.tag_config("start", background="green", 
                            foreground="white") 
    
    else:
        # generowanie pdfow z protokolow .htm
        for idx, file in enumerate(paths_tuple):
            pdfkit.from_file(
        file,
        f"pdfs/{list_names_files[idx]}.pdf",
        # verbose=True,
        options={"enable-local-file-access": True},
    )
            # add tag using indices for the 
            # part of text to be highlighted
            text.insert(f"{idx+1}.0", "OK! ") 
            text.tag_add("start", f"{idx+1}.0", f"{idx+1}.3") 
            #configuring a tag called start 
            text.tag_config("start", background="green", 
                            foreground="white")      
                
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
        idx_val += 1
                  
    # os.startfile("Badanie_.txt", "print")

# Utworzenie okna aplikacji 
root = Tk()
root.title("Simple print Example")
root.geometry("800x500")
# l = Label(text="Select ur desired file to print", bg="gray")
# l.pack(fill=X)
frame_button = Frame(root)
frame_button.pack(pady=5)
frame_entry = Frame(root)
frame_entry.pack(pady=5)
frame_bar = Frame(root)
frame_bar.pack(pady=5)
frame_text = Frame(root, relief="groove")
frame_text.pack(pady=5)

btn_select = Button(frame_button,
    text="Select File", width=15, bg="dark green", command=select_files
)
btn = Button(frame_button, text="createPDF", width=15,
             command=thread_to_bar, state="disabled")


btn_select.grid(padx=5, pady=5, column=0, row=0)
btn.grid(padx=5 ,pady=5, column=1, row=0)

info_lab = Label(frame_text, font=("Arial Black", "10"), fg="blue", text="No files")
text = Text(frame_text, wrap='none', bg='black', fg="white")

info_lab.grid(padx=5, pady=5, row=0, column=0)
text.grid(padx=5, column=0, row=1)



root.mainloop()