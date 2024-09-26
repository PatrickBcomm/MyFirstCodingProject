import pdfminer
from pdfminer.high_level import extract_text
import re
import pandas as pd

# Function to extract sales figures from the PDF
def extract_sales_from_pdf(pdf_file):
    # Extract text from the PDF
    text = extract_text(pdf_file)

    # Search for sales figures on page 6
    sales_pattern = r"Sales\s*€([0-9,.]+)\s*Bn"
    match = re.search(sales_pattern, text)
    
    if match:
        sales_figure = match.group(1)
        return sales_figure
    else:
        return "Sales figure not found"

# Function to export the sales figure to Excel
def export_to_excel(sales_figure, output_file):
    # Create a DataFrame to store the sales figure
    df = pd.DataFrame({"Sales Figure (Bn)": [sales_figure]})
    
    # Export to Excel
    df.to_excel(output_file, index=False)
    print(f"Sales figure exported to {output_file}")

# Main function to handle the file upload and export
def main():
    pdf_file = '/path/to/your/pdf/file.pdf'  # Replace with the path to your PDF file
    output_file = 'loreal_sales.xlsx'  # Output Excel file
    
    sales_figure = extract_sales_from_pdf(pdf_file)
    print(f"Extracted Sales Figure: €{sales_figure} Bn")
    
    # Export to Excel
    export_to_excel(sales_figure, output_file)

if __name__ == "__main__":
    main()
