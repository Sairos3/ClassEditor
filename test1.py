import re
import PyPDF2
from tkinter import filedialog, Tk

def select_pdf_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(title="Select PDF file", filetypes=[("PDF Files", "*.pdf")])

def extract_text_from_pdf(pdf_file):
    full_text = ""
    with open(pdf_file, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
    return full_text

def main():
    pdf_file = select_pdf_file()
    if not pdf_file:
        print("No PDF file selected. Exiting...")
        return

    text_content = extract_text_from_pdf(pdf_file)
    with open("schedule.txt", "w", encoding="utf-8") as txt_file:
        txt_file.write(text_content)

if __name__ == "__main__":
    main()
