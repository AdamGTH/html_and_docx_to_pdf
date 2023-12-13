from PyPDF2 import PdfWriter, PdfReader

inputpdf = PdfReader(open("pdfs/out0.pdf", "rb"))
inputpdf2 = PdfReader(open("pdfs/out1.pdf", "rb"))
# for i in range(len(inputpdf.pages)):
#     output = PdfWriter()
#     output.add_page(inputpdf.pages[i])
#     with open("document-page%s.pdf" % i, "wb") as outputStream:
#         output.write(outputStream)

output = PdfWriter()
output.add_page(inputpdf.pages[0])
output.add_page(inputpdf2.pages[0])

with open("doc.pdf", "wb") as out_stream:
    output.write(out_stream)