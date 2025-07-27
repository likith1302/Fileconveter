from pdf2docx import Converter
from docx2pdf import convert
import tkinter as tk
from tkinter import filedialog
from PIL import Image
from fpdf import FPDF
import pandas as pd
import os

# 🧰 Format Conversion Functions

def csv_to_excel(csv_file, excel_file):
    df = pd.read_csv(csv_file)
    df.to_excel(excel_file, index=False)
    print(f"[✓] CSV converted successfully → {excel_file}")

def excel_to_json(excel_file, json_file):
    df = pd.read_excel(excel_file)
    df.to_json(json_file, orient="records", indent=4)
    print(f"[✓] Excel converted successfully → {json_file}")

def text_to_pdf(txt_file, pdf_file):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    with open(txt_file, 'r') as file:
        for line in file:
            pdf.cell(200, 10, txt=line.strip(), ln=True)

    pdf.output(pdf_file)
    print(f"[✓] Text file converted → {pdf_file}")

def image_to_pdf(image_path, pdf_path):
    image = Image.open(image_path)
    image.convert("RGB").save(pdf_path)
    print(f"[✓] Image converted → {pdf_path}")

def convert_pdf_to_word(pdf_path, docx_path, start=0, end=None):
    print(f"🔄 Converting PDF → Word...\nInput: {pdf_path}\nOutput: {docx_path}")
    converter = Converter(pdf_path)
    converter.convert(docx_path, start=start, end=end)
    converter.close()
    print(f"[✓] PDF successfully converted to Word → {docx_path}")

def convert_word_to_pdf(word_file):
    convert(word_file)
    output = word_file.replace('.docx', '.pdf')
    print(f"[✓] Word file converted → {output}")

# 📁 Reusable File Picker
def pick_file(file_types, prompt="Select a file"):
    root = tk.Tk()
    root.withdraw()
    print(f" {prompt}")
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if not file_path:
        print("[!] No file selected.")
    return file_path

# 🧭 Main Menu
def main():
    options = {
        "1": "Convert PDF to Word",
        "2": "Convert Word to PDF",
        "3": "Convert Image to PDF",
        "4": "Convert Text File to PDF",
        "5": "Convert CSV to Excel",
        "6": "Convert Excel to JSON",
        "0": "Exit"
    }

    while True:
        print("\n📌 Format Converter Menu")
        for key, value in options.items():
            print(f" {key}. {value}")
        choice = input("📝 Enter your choice: ")

        if choice == "1":
            pdf = pick_file([("PDF Files", "*.pdf")], "Choose a PDF file to convert:")
            if pdf:
                docx_path = os.path.splitext(pdf)[0] + "_converted.docx"
                convert_pdf_to_word(pdf, docx_path)

        elif choice == "2":
            docx = pick_file([("Word Files", "*.docx")], "Choose a Word file to convert:")
            if docx:
                convert_word_to_pdf(docx)

        elif choice == "3":
            img = pick_file([("Image Files", "*.png *.jpg *.jpeg *.bmp")], "Choose an image file:")
            if img:
                pdf_path = os.path.splitext(img)[0] + ".pdf"
                image_to_pdf(img, pdf_path)

        elif choice == "4":
            txt = pick_file([("Text Files", "*.txt")], "Choose a text file:")
            if txt:
                pdf_path = os.path.splitext(txt)[0] + ".pdf"
                text_to_pdf(txt, pdf_path)

        elif choice == "5":
            csv = pick_file([("CSV Files", "*.csv")], "Choose a CSV file:")
            if csv:
                excel_path = os.path.splitext(csv)[0] + ".xlsx"
                csv_to_excel(csv, excel_path)

        elif choice == "6":
            excel = pick_file([("Excel Files", "*.xlsx")], "Choose an Excel file:")
            if excel:
                json_path = os.path.splitext(excel)[0] + ".json"
                excel_to_json(excel, json_path)

        elif choice == "0":
            print("👋 Thanks for using the converter! Goodbye.")
            break

        else:
            print("[!] Invalid choice. Please select a valid option.")

# 🔁 Entry point
if __name__ == "__main__":
    main()