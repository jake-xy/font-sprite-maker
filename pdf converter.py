import os
from docx2pdf import convert

print("converting it to pdf...")
convert("out.docx", "out.pdf")

print("done")
os.system("pause > nul")