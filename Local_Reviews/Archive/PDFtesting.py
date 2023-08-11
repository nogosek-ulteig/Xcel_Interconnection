import pdftotext

file_path = r"G:\2021\21.00016\Reviews\JN\IA30816 - JN\DAVID BREZINSKI LINE DIAGRAM - 43862.pdf"

pdf = pdftotext.PDF(file_path)

for page in pdf:
    print(page)
