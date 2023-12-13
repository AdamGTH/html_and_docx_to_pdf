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

def progress_bar(actual_iterable):
    global ilosc
    pcnt = actual_iterable * 100 / ilosc
    bar["value"] = pcnt
    bar_val["text"] = f'{round(bar["value"],2)} %'
    if bar["value"] == 100.0:
       bar_val["text"] = "Is finished !!!"

def select_files():
    global paths_tuple, path_this, ilosc, list_names_files
    paths_tuple = askopenfilenames(
        parent=root, filetypes=[
            ("Text Files", "*.htm"), ("Text Files", "*.html")]
    )
    # print(paths_tuple)
    path_this = os.getcwd()
    ilosc = len(paths_tuple)
    list_names_files = [name[name.find("Badanie"):] for name in paths_tuple]
    btn.configure(state="normal")
    info_lab.configure(text=f"loaded {len(paths_tuple)} files")
    

    for idx, val in enumerate(paths_tuple):
        text.insert(f"{idx}.0", val[val.find(
            "Badanie"):val.find("Badanie")+40] + "...\n")
    
     
   
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
            if "path_to_template" in line:
                path_to_template = line.split("=")[1].strip()
    
    # generowanie potwierdzienia w .docx
    out_file_doc = "POTWIERDZENIE_WZORCOWANIA-%d.%.2d.%.2d" % (current_time.year, current_time.month, current_time.day)
    doc_name_docx = f"{out_file_doc}.docx"
    doc_name_pdf = f"{path_to_save}{out_file_doc}.pdf"
    doc = DocxTemplate("szablon_potw.docx")
    context = {"nr_zgloszenia": nr_zgl.get(), "nr_oum": nr_oum.get(), "data": data.get(), "ilosc": ilosc_zbadanych.get(), "osoba_emiter": emit_man.get(), "osoba_oum": oum_man.get()}
    doc.render(context)
    doc.save(doc_name_docx)
    
    # generowanie pdfa z pliku potwierdzenia generated_doc.docx za pomoca comendy konsoli, dziala tylko z Libre office:
    # os.system(f"\"C:/Program Files/LibreOffice/program/swriter.exe\" --headless --convert-to pdf --outdir {path_to_save} {doc_name_docx}")
    commandStrings = [path_to_libre_engine, "--headless", "--convert-to", "pdf", "--outdir", path_to_save, doc_name_docx]
    # print(" ".join(commandStrings))
    retCode = subprocess.call(commandStrings)
    
    
    # generowanie pdfow z protokolow .htm
    for idx, file in enumerate(paths_tuple):
        pdfkit.from_file(
    file,
    f"pdfs/out{idx}.pdf",
    # verbose=True,
    options={"enable-local-file-access": True},
)
        progress_bar(idx)
        idx_val = idx
    
              
        
    # generowanie zestawienia w pliku .txt i utworzenie pdfa z pliku .txt
    file_name_apply = "WYKAZ NR. PRZEKŁ.PRĄD. - WZORCOWANIE %d.%.2d.%.2d.txt" % (current_time.year, current_time.month, current_time.day)
    
    with open(f"{path_to_save}{file_name_apply}", "w+", encoding="utf8") as f:
        f.write(f"Przebadano {len(list_names_files)} szt\n\n")
        for line in list_names_files:
            f.write(line+"\n")
    
    pdfkit.from_file(f"{path_to_save}{file_name_apply}", "pdfs/Potwierdzenie.pdf", options={"enable-local-file-access": True, "encoding": "utf8"})
    
    # generowanie zbiorczego pliku pdf
    pdfs_reader = []
    for idx in range(len(paths_tuple)):
        pdfs_reader.append(PdfReader(open(f"pdfs/out{idx}.pdf", "rb")))
    pdfs_reader.append(PdfReader(open("pdfs/Potwierdzenie.pdf", "rb")))
    pdfs_reader.append(PdfReader(open(doc_name_pdf, "rb")))
        
    output_pdf = PdfWriter()
    
    for idx, file in enumerate(pdfs_reader):
        for page in file.pages:
            output_pdf.add_page(page)
        
    with open("pdfs/all_doc.pdf", "wb") as out_stream:
        output_pdf.write(out_stream)
        idx_val += 1
    
                
    progress_bar(idx_val)
        
    # os.startfile("Badanie_.txt", "print")

# Utworzenie okna aplikacji 
root = Tk()
root.title("Simple print Example")
root.geometry("600x500")
# l = Label(text="Select ur desired file to print", bg="gray")
# l.pack(fill=X)
frame_button = Frame(root)
frame_button.pack(pady=5)
frame_entry = Frame(root)
frame_entry.pack(pady=5)
frame_bar = Frame(root)
frame_bar.pack(pady=5)
frame_text = Frame(root)
frame_text.pack(pady=5)

btn_select = Button(frame_button,
    text="Select File", width=15, bg="dark green", command=select_files
)
btn = Button(frame_button, text="createPDF", width=15,
             command=thread_to_bar, state="disabled")


btn_select.grid(padx=5, pady=5, column=0, row=0)
btn.grid(padx=5 ,pady=5, column=1, row=0)

nr_oum = Entry(frame_entry, width=30)
nr_oum.insert(0, "W3/---/%d" % (current_time.year))
nr_zgl = Entry(frame_entry, width=30)

nr_zgl.insert(0, "Z/%d/%d/%.2d" % (current_time.year,current_time.month,current_time.day-1))
oum_man = ttk.Combobox(frame_entry, state="readonly", values=oum_mans, width=30)
emit_man = ttk.Combobox(frame_entry, state="readonly", values=["Adam Masiarz", "Arkadiusz Świerczek", "Marek Lampa"], width=30)
ilosc_zbadanych = Entry(frame_entry, width=30)
ilosc_zbadanych.insert(0, "200 szt.")
data = Entry(frame_entry, width=30)
data.insert(0, f"%.2d.%d.%d r." % (current_time.day,current_time.month, current_time.year))

nr_oum_label = Label(frame_entry, text="NR OUM:")
nr_zgl_label = Label(frame_entry, text="NR ZGŁ.:")
oum_man_label = Label(frame_entry, text="SPRAWDZIŁ:")
emit_man_label = Label(frame_entry, text="WYKONAŁ:")
ilosc_zbadanych_label = Label(frame_entry, text="ILOŚĆ:")
data_label = Label(frame_entry, text="DATA:")

nr_oum_label.grid(pady=5, column=0, row=0)
nr_oum.grid(pady=5, column=1, row=0)
nr_zgl_label.grid(pady=5, column=0, row=1)
nr_zgl.grid(pady=5, column=1, row=1)
data_label.grid(pady=5, column=0, row=2)
data.grid(pady=5, column=1, row=2)
oum_man_label.grid(pady=5, column=0, row=3)
oum_man.grid(pady=5, column=1, row=3)
emit_man_label.grid(pady=5, column=0, row=4)
emit_man.grid(pady=5, column=1, row=4)
ilosc_zbadanych_label.grid(pady=5, column=0, row=5)
ilosc_zbadanych.grid(pady=5, column=1, row=5)

bar_val = Label(frame_bar, text="0.0 %")
bar = ttk.Progressbar(
    frame_bar,
    orient="horizontal",
    mode="determinate",
    length=300,
)

bar_val.grid(padx=5, pady=5, column=0, row=0)
bar.grid(padx=10, pady=10, column=0, row=1)

info_lab = Label(frame_text, font=("Arial Black", "10"), fg="blue", text="No files")
text = ScrolledText(frame_text, height=10, width=50)

info_lab.grid(padx=5, pady=5, row=0, column=0)
text.grid(padx=5, column=0, row=1)



root.mainloop()