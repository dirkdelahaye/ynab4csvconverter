##########################
## YNAB 4 CSV Convertor ##
## Dirk Delahaye, 2023  ##
##########################
import pandas as pd
import sys
import os
import warnings
from datetime import datetime

# Import necessary libraries for PDF processing
import pdfplumber
import re

# Suppress the specific warning related to openpyxl default style
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Function to process XLSX file
def process_xlsx(xlsx_filename):
    try:
        df = pd.read_excel(xlsx_filename)

        # Convert the "Date" column to the desired format "DD/MM/YYYY"
        df['Date'] = df['Verrichtingsdatum'].apply(lambda x: x.strftime('%d/%m/%Y'))

        # Create a new DataFrame with the mapped columns
        ynab_df = pd.DataFrame({
            'Account': 'Argenta',
            'Flag': '',
            'Date': df['Date'],
            'Payee': df['Naam tegenpartij'],
            'Category': '',
            'Master Category': '',
            'Sub Category': '',
            'Memo': df['Mededeling'],
            'Outflow': df['Bedrag'].apply(lambda x: -x if x < 0 else 0),
            'Inflow': df['Bedrag'].apply(lambda x: x if x >= 0 else 0),
            'Cleared': ''
        })

        # Form the output filename
        original_filename = os.path.splitext(os.path.basename(xlsx_filename))[0]
        output_filename = f"ynab4_import_{original_filename}.csv"

        # Save the transformed data as YNAB4 CSV
        ynab_df.to_csv(output_filename, index=False)
        print(f"CSV file saved as {output_filename}")

    except Exception as e:
        print(f"Error processing the Excel file: {e}")

# Function to process PDF file
def process_pdf_statement(pdf_filename):
    try:
        with pdfplumber.open(pdf_filename) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text()

        lines = text.splitlines()

        # Initialize variables
        start_processing = False

        # Initialize lists to store extracted data
        dates = []
        payees = []
        outflows = []
        inflows = []

        # Define a regular expression pattern for transaction lines
        transaction_pattern = r'^(\d{2}/\d{2}) (\d{2}/\d{2}) (.+?) ([\d.,]+)([+-])'
        
        for line in lines:
            print("Evaluating " + line)
            if start_processing and not line.strip():
                # Stop processing when encountering an empty line after starting
                break
            elif not start_processing and line.startswith("TRAD NA ST AU CM TIE VERD RA ET KU EM NING OMSCHRIJVING BEDRAG (EUR)"):
                # Start processing when encountering the header line
                start_processing = True
            elif start_processing and re.fullmatch(transaction_pattern, line):
                print("Full Match.")
                match = re.fullmatch(transaction_pattern, line)
                if match:
                    date, _, payee, amount, sign = match.groups()
                    print("Amount=" + amount)
                    dates.append(date)
                    payees.append(payee.strip())

                    # Replace thousands separators and replace comma with period as the decimal separator
                    amount = float(amount.replace('.', '').replace(',', '.'))
                    if line.endswith('-'):
                        outflows.append(amount)
                        inflows.append(0.0)
                    else:
                        outflows.append(0.0)
                        inflows.append(amount)
                else: 
                    print("No Match.")
                    
        # Create a new DataFrame for PDF data
        pdf_df = pd.DataFrame({
            'Account': 'Argenta Mastercard',
            'Flag': '',
            'Date': dates,
            'Payee': payees,
            'Category': '',
            'Master Category': '',
            'Sub Category': '',
            'Memo': [''] * len(dates),
            'Outflow': outflows,
            'Inflow': inflows,
            'Cleared': ''
        })


        # Form the output filename for PDF data
         # Form the output filename
        pdf_original_filename = os.path.splitext(os.path.basename(pdf_filename))[0]
        pdf_output_filename = f"ynab4_import_pdf_{pdf_original_filename}.csv"
        #output_filename = f"ynab4_import_{original_filename}.csv"

        # Save the PDF data as YNAB4 CSV
        pdf_df.to_csv(pdf_output_filename, index=False)
        print(f"PDF data saved as {pdf_output_filename}")

    except Exception as e:
        print(f"Error processing the PDF file: {e}")
        return [], [], [], []

# Check the number of command line arguments
if len(sys.argv) != 2:
    print("Usage: python script.py <filename>")
    sys.exit(1)

# Get the input filename from the command line argument
input_filename = sys.argv[1]

# Check the file extension to determine the type of file
file_extension = os.path.splitext(input_filename)[1].lower()

if file_extension == ".xlsx":
    process_xlsx(input_filename)
elif file_extension == ".pdf":
    process_pdf_statement(input_filename)
else:
    print("Unsupported file format. Supported formats: .xlsx, .pdf")
