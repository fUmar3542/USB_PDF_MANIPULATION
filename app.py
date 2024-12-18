import fitz
import csv
import os
import re
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import io

mode = '0'


def log_error(exception):
    """Log exceptions to errors.txt file."""
    with open("errors.txt", "a") as error_file:
        error_file.write(f"{exception}\n")


def add_text_to_pdf(input_pdf_path, output_pdf_path, data, font_size=15):
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for i, page in enumerate(reader.pages):
            packet = io.BytesIO()
            c = canvas.Canvas(packet)

            if data[i][2] == 0:
                c.setFont("Helvetica", font_size)
                c.drawString(10, 260, data[i][0])

                if data[i][1] != '1':
                    c.setFont("Helvetica", 40)
                    c.drawString(170, 340, data[i][1])
            else:
                global mode
                c.setFont("Helvetica", 12)
                if mode == "1":
                    c.drawString(20, 113, data[i][0])
                else:
                    c.drawString(20, 113, "")

                if data[i][1] != '1':
                    c.setFont("Helvetica", 40)
                    c.drawString(220, 30, data[i][1])

            c.save()

            packet.seek(0)
            overlay_pdf = PdfReader(packet)
            page.merge_page(overlay_pdf.pages[0])
            writer.add_page(page)

        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)

        print("\nProcess completed successfully...")
    except Exception as ex:
        log_error(ex)
        print("There is some error occurred during writing data to pdf...")


def read_excel(input_csv_path):
    column_values = {}
    try:
        with open(input_csv_path, 'r', newline='', errors='ignore') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                try:
                    value_t = row['quantity-purchased']
                    value_k = row['reference2']
                    value_n = row['Order Quantity']
                    column_values[value_t] = [value_k, value_n]
                except Exception as ex:
                    log_error(ex)
    except Exception as ex:
        log_error(ex)
        print("There is some error occurred during processing...")
    finally:
        return column_values


def compare_with_excel(excel_dict, values):
    numbers = []
    try:
        keys = excel_dict.keys()
        for i in range(len(values[0])):
            x = values[0][i]
            if x in keys:
                numbers.append(excel_dict[x])
            else:
                numbers.append(["", ""])
            if values[1][i] == 0:
                numbers[-1][0] = numbers[-1][0].replace('-FX', '')
            numbers[-1] += [values[1][i]]
    except Exception as ex:
        log_error(ex)
        print("There is some error occurred during processing...")
    finally:
        return numbers


def ocr_file(pdf_path):
    values = [[], []]
    try:
        pdf_document = fitz.open(pdf_path)
        for page_number in range(pdf_document.page_count):
            try:
                page = pdf_document[page_number]
                text = page.get_text()
                pattern = r'(.*?)\nBusiness Account'
                match = re.search(pattern, text, re.DOTALL)
                if match:
                    pattern = r'\n([^(\n)]*)$'
                    match = re.search(pattern, match.group(1).strip(), re.DOTALL)
                    if match:
                        val = match.group(1).strip()
                        val = val.replace('-', '')
                        values[0].append(val)
                        values[1].append(0)
                else:
                    pattern = r'Customer Ref:\s*(\d+)\s*/\s*([0-9A-Za-z-]+)'
                    match = re.search(pattern, text)
                    if match:
                        values[0].append(match.group(2).strip())
                        values[1].append(1)
                    else:
                        values[0].append("")
                        values[1].append(0)
            except Exception as ex:
                log_error(ex)
                values[0].append("")
                values[1].append(0)
        pdf_document.close()
    except Exception as ex:
        log_error(ex)
        print("There is some error occurred during processing...")
    finally:
        return values


def read_mode_file(file_path):
    try:
        with open(file_path, "r") as file:
            first_char = file.read(1)
            if first_char in ('0', '1'):
                return first_char
            else:
                print("Invalid content in file. Expected '0' or '1'.")
                return None
    except FileNotFoundError:
        log_error(f"File '{file_path}' not found.")
        return None
    except Exception as ex:
        log_error(ex)
        return None


def main():
    try:
        mode1 = read_mode_file(file_path="mode.txt")
        if mode1 is not None:
            global mode
            mode = mode1
            input_pdf = "input.pdf"
            output_pdf = "output.pdf"
            excel_path = "input.csv"

            if not os.path.exists(input_pdf):
                print(f"The file '{input_pdf}' does not exist in the current folder.")
                return
            if not os.path.exists(excel_path):
                print(f"The file '{excel_path}' does not exist in the current folder.")
                return

            result_dict = read_excel(excel_path)
            if not result_dict:
                print("No data found in respective columns in csv file")
                return

            values = ocr_file(input_pdf)
            if not values[0]:
                print("No label number found in the pdf file")
                return

            data = compare_with_excel(result_dict, values)
            if not data:
                print("No corresponding data found in csv file")
                return

            add_text_to_pdf(input_pdf, output_pdf, data)
    except Exception as ex:
        log_error(ex)
        print("There is some error occurred during processing...")


main()
