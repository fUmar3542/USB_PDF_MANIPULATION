import fitz  # PyMuPDF
import csv
import os
import re


def add_text_to_pdf(input_pdf_path, output_pdf_path, data, font_size=20):
    try:
        # Open the PDF file
        pdf_document = fitz.open(input_pdf_path)
        for i in range(pdf_document.page_count):
            page = pdf_document[i]
            # Define the position for top-left corner (in points)
            position = (10, 155)
            # Create a new TextPage object
            text_page = page.insert_text(position, data[i], fontname="helv", fontsize=font_size)
        # Save the changes
        pdf_document.save(output_pdf_path)

        # Close the PDF document
        pdf_document.close()
    except Exception as ex:
        print("There is some error occurred during writing data to pdf...")
        print(ex)


def read_excel(input_csv_path):
    column_values = {}
    try:
        # Open the CSV file
        with open(input_csv_path, 'r', newline='') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                value_t = row['quantity-purchased']
                value_k = row['reference2']

                # Store values in the dictionary
                column_values[value_t] = value_k
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)
    finally:
        return column_values


def compare_with_excel(excel_dict, values):
    numbers = []
    try:
        keys = excel_dict.keys()
        for x in values:
            if x in keys:
                numbers.append(excel_dict[x])
            else:
                numbers.append("")
            numbers[-1] = numbers[-1].replace('-FX', '')
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)
    finally:
        return numbers


def ocr_file(pdf_path):
    values = []
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
                        values.append(val)
                else:
                    values.append("")
            except:
                values.append("")
                pass
        pdf_document.close()
    except Exception as ex:
        print("There is some error occurred during processing...")
        print(ex)
    finally:
        return values


def main():
    try:
        # Example usage:
        input_pdf = "input.pdf"
        output_pdf = "output.pdf"
        excel_path = "input.csv"  # Replace with the path to your Excel file
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
            print("No data found in column T and K in csv file")
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
