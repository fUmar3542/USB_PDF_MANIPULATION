import fitz
import csv
import os
import re
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import io

mode = '0'


def add_text_to_pdf(input_pdf_path, output_pdf_path, data, font_size=15):
    try:
        # Read the existing PDF
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for i, page in enumerate(reader.pages):
            # Create a canvas for the overlay
            packet = io.BytesIO()
            c = canvas.Canvas(packet)

            # Add text based on data
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

            # Merge the overlay with the existing page
            packet.seek(0)
            overlay_pdf = PdfReader(packet)
            page.merge_page(overlay_pdf.pages[0])
            writer.add_page(page)

        # Write to the output file
        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)

        print("\nProcess completed successfully...")
    except Exception as ex:
        print("There is some error occurred during writing data to pdf...")
        print(ex)


def read_excel(input_csv_path):
    column_values = {}
    try:
        # Open the CSV file
        with open(input_csv_path, 'r', newline='', errors='ignore') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                try:
                    value_t = row['quantity-purchased']
                    value_k = row['reference2']
                    value_n = row['Order Quantity']

                    # Store values in the dictionary
                    column_values[value_t] = [value_k, value_n]
                except:
                    pass
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)
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
        print("There is some error occurred during processing...")
        print(ex)
    finally:
        return numbers


def ocr_file(pdf_path):
    values = [[], []]
    try:
        # Open the PDF file
        pdf_document = fitz.open(pdf_path)
        # Iterate through the pages
        for page_number in range(pdf_document.page_count):
            try:
                page = pdf_document[page_number]
                # Extract text from the page
                text = page.get_text()
                # Define a regular expression pattern
                pattern = r'(.*?)\nBusiness Account'
                # Use re.search to find the match
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
                values[0].append("")
                values[1].append(0)
                pass
        pdf_document.close()
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)
    finally:
        return values


def read_mode_file(file_path):
    try:
        with open(file_path, "r") as file:
            first_char = file.read(1)  # Read the first character only
            if first_char in ('0', '1'):
                return first_char
            else:
                print("Invalid content in file. Expected '0' or '1'.")
                return None
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None


def main():
    try:
        mode1 = read_mode_file(file_path="mode.txt")
        if mode1 is not None:
            global mode
            mode = mode1
            input_pdf = "input.pdf"
            output_pdf = "output.pdf"
            excel_path = "input.csv"    # Replace with the path to your Excel file
            # Check if the file exists
            if not os.path.exists(input_pdf):
                print(f"The file '{input_pdf}' does not exist in the current folder.")
                return
            if not os.path.exists(excel_path):
                print(f"The file '{excel_path}' does not exist in the current folder.")
                return

            # Get values from columns T and K as a dictionary
            result_dict = read_excel(excel_path)
            if not result_dict:
                print("No data found in respective columns in csv file")
                return

            values = ocr_file(input_pdf)
            if not result_dict:
                print("No label number found in the pdf file")
                return

            data = compare_with_excel(result_dict, values)
            if not result_dict:
                print("No corresponding data found in csv file")
                return

            add_text_to_pdf(input_pdf, output_pdf, data)
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)


main()
