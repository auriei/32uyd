import pdfplumber
import os
import pandas as pd
import re
from datetime import datetime

# Configuration (Consider moving to a config file or class)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Gets the directory of the current script
LOG_DIR = os.path.join(BASE_DIR, '..', '..', 'logs')  # Path to the logs directory: /app/logs
TEMP_DIR = os.path.join(BASE_DIR, '..', '..', 'temp') # Path to the temp directory: /app/temp
DATA_DIR = os.path.join(BASE_DIR, '..', '..', 'data') # Path to the data directory: /app/data
QC_DIR = os.path.join(DATA_DIR, 'qc') # Path to the QC directory: /app/data/qc
REPORTS_DIR = os.path.join(QC_DIR, 'reports') # Path to the reports directory: /app/data/qc/reports
FORMS_DIR = os.path.join(QC_DIR, 'forms') # Path to the forms directory: /app/data/qc/forms


# Ensure directories exist
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs(FORMS_DIR, exist_ok=True)


# Logging setup (Basic)
LOG_FILE = os.path.join(LOG_DIR, 'pdf_processing.log')

def log_message(message):
    """Appends a timestamped message to the log file."""
    with open(LOG_FILE, 'a') as f:
        f.write(f"{datetime.now()}: {message}\n")

log_message("Script execution started.")


class PDFProcessor:
    def __init__(self, pdf_path):
        if not os.path.exists(pdf_path):
            log_message(f"ERROR: PDF file not found at {pdf_path}")
            raise FileNotFoundError(f"PDF file not found at {pdf_path}")
        self.pdf_path = pdf_path
        self.text_content = ""
        self.tables = []
        log_message(f"Initialized PDFProcessor with {pdf_path}")

    def extract_text(self):
        """Extracts all text from the PDF."""
        log_message(f"Starting text extraction for {self.pdf_path}")
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                full_text = []
                for i, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        full_text.append(f"--- Page {i+1} ---\n{page_text}")
                    else:
                        full_text.append(f"--- Page {i+1} ---\n[No text extracted]")
                        log_message(f"No text extracted from page {i+1} of {self.pdf_path}")
                self.text_content = "\n".join(full_text)
            log_message(f"Successfully extracted text from {self.pdf_path}")
            return self.text_content
        except Exception as e:
            log_message(f"ERROR: Failed to extract text from {self.pdf_path}. Error: {e}")
            # Optionally, re-raise the exception if you want the script to stop
            # raise
            return None # Or return an empty string, or some error indicator


    def extract_tables(self):
        """Extracts all tables from the PDF into pandas DataFrames."""
        log_message(f"Starting table extraction for {self.pdf_path}")
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_tables = page.extract_tables()
                    if page_tables:
                        log_message(f"Found {len(page_tables)} table(s) on page {i+1} of {self.pdf_path}")
                        for table_num, table_data in enumerate(page_tables):
                            df = pd.DataFrame(table_data[1:], columns=table_data[0]) # Assumes first row is header
                            self.tables.append({"page": i + 1, "table_num": table_num + 1, "dataframe": df})
                    else:
                        log_message(f"No tables found on page {i+1} of {self.pdf_path}")
            log_message(f"Successfully extracted {len(self.tables)} tables from {self.pdf_path}")
            return self.tables
        except Exception as e:
            log_message(f"ERROR: Failed to extract tables from {self.pdf_path}. Error: {e}")
            # Optionally, re-raise
            return [] # Return empty list on failure

    def save_text(self, output_path=None):
        """Saves the extracted text to a file."""
        if not self.text_content:
            log_message("No text content to save.")
            print("No text content to save. Extract text first.")
            return

        if output_path is None:
            base, ext = os.path.splitext(os.path.basename(self.pdf_path))
            output_path = os.path.join(TEMP_DIR, f"{base}_extracted_text.txt")
        
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(self.text_content)
            log_message(f"Text content saved to {output_path}")
            print(f"Text content saved to {output_path}")
        except IOError as e:
            log_message(f"ERROR: Failed to save text content to {output_path}. Error: {e}")
            print(f"Error saving text to {output_path}: {e}")


    def save_tables(self, output_dir=None, format="csv"):
        """Saves extracted tables to files (CSV or Excel)."""
        if not self.tables:
            log_message("No tables to save.")
            print("No tables to save. Extract tables first.")
            return

        if output_dir is None:
            output_dir = os.path.join(TEMP_DIR, f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_tables")
        
        os.makedirs(output_dir, exist_ok=True)

        for table_info in self.tables:
            page_num = table_info['page']
            table_idx = table_info['table_num']
            df = table_info['dataframe']
            
            try:
                if format == "csv":
                    filename = os.path.join(output_dir, f"page_{page_num}_table_{table_idx}.csv")
                    df.to_csv(filename, index=False)
                elif format == "excel":
                    filename = os.path.join(output_dir, f"page_{page_num}_table_{table_idx}.xlsx")
                    df.to_excel(filename, index=False)
                else:
                    log_message(f"Unsupported table format {format}. Supported formats: 'csv', 'excel'.")
                    print(f"Unsupported table format: {format}")
                    return # Or skip this table
                log_message(f"Table page {page_num}, table {table_idx} saved to {filename}")
                print(f"Table page {page_num}, table {table_idx} saved to {filename}")
            except Exception as e: # More specific exceptions can be caught if needed
                log_message(f"ERROR: Failed to save table (Page {page_num}, Table {table_idx}) to {filename}. Error: {e}")
                print(f"Error saving table (Page {page_num}, Table {table_idx}) to {filename}: {e}")


# Example Usage
def main():
    """Main function to demonstrate PDFProcessor usage."""
    log_message("Main function execution started.")
    # Use a PDF from the /app/ directory for testing
    # Pick one of the existing PDF files in the root for this example.
    # This requires knowing a PDF filename that will be present in the /app directory when run.
    # Let's assume 'EB05曲轴OP110.PDF' is available at /app/ as per previous listings.
    
    # Constructing path relative to BASE_DIR (src/core) to root /app
    # ../.. goes from src/core to src, then to /app
    pdf_root_dir = os.path.join(BASE_DIR, '..', '..') 
    
    # List of available PDFs in the root directory (from previous observations)
    # This is a placeholder. In a real scenario, you might get this list dynamically or from a config.
    available_pdfs = [
        "EB05曲轴OP110.PDF", "EB05曲轴OP20A.PDF", "EB05曲轴OP30.PDF", 
        "EB05曲轴OP50.PDF", "EB05曲轴OP70.PDF", "EB05曲轴OP90.PDF"
        # Add more known PDF names if needed
    ]

    if not available_pdfs:
        log_message("No example PDFs specified in the script to process.")
        print("Please specify an example PDF filename in the script for testing.")
        return

    # Using the first PDF from the list for the example
    example_pdf_filename = available_pdfs[0]
    example_pdf_path = os.path.join(pdf_root_dir, example_pdf_filename) 
    
    print(f"Attempting to process PDF: {example_pdf_path}")
    log_message(f"Attempting to process PDF: {example_pdf_path}")

    if not os.path.exists(example_pdf_path):
        log_message(f"CRITICAL: Example PDF '{example_pdf_filename}' not found at expected path '{example_pdf_path}'. Make sure it's in the /app directory.")
        print(f"CRITICAL: Example PDF '{example_pdf_filename}' not found at expected path '{example_pdf_path}'. Cannot run example.")
        # Try to list files in pdf_root_dir to help debug if file is missing
        try:
            print(f"Files in {pdf_root_dir}: {os.listdir(pdf_root_dir)}")
        except Exception as e:
            print(f"Could not list files in {pdf_root_dir}: {e}")
        return

    processor = PDFProcessor(example_pdf_path)

    # Extract and save text
    text = processor.extract_text()
    if text:
        processor.save_text() # Saves to temp directory by default

    # Extract and save tables
    tables = processor.extract_tables()
    if tables:
        processor.save_tables(format="csv") # Saves to temp directory by default as CSV
        # processor.save_tables(format="excel") # Optionally save as Excel

    log_message("Main function execution finished.")
    print("Processing finished. Check logs and temp directories for output.")

if __name__ == "__main__":
    main()
