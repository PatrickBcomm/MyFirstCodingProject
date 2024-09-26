from flask import Flask, render_template, request, redirect, send_file
import os
from pdfminer.high_level import extract_text
import re
import pandas as pd

# Initialize the Flask application
app = Flask(__name__)

# Folder to store uploaded files
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Function to extract sales figures from the PDF
def extract_sales_from_pdf(pdf_file):
    # Extract text from the PDF file
    text = extract_text(pdf_file)
    
    # Regular expression to search for sales figure (e.g., "Sales €38.26 Bn")
    sales_pattern = r"Sales\s*€([0-9,.]+)\s*Bn"
    match = re.search(sales_pattern, text)

    if match:
        # Return the matched sales figure
        return match.group(1)
    else:
        return "Sales figure not found"

# Function to export the sales figure to an Excel file
def export_to_excel(sales_figure, output_file):
    # Create a DataFrame with the sales figure
    df = pd.DataFrame({"Sales Figure (Bn)": [sales_figure]})

    # Export the DataFrame to an Excel file
    df.to_excel(output_file, index=False)

# Home route to display the HTML form
@app.route('/')
def index():
    return render_template('financial data.html')

# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if a file is included in the request
    if 'file' not in request.files:
        return redirect(request.url)

    file = request.files['file']

    # Check if a file was actually uploaded
    if file.filename == '':
        return redirect(request.url)

    if file:
        # Save the uploaded file to the uploads directory
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        # Extract sales figure from the uploaded PDF
        sales_figure = extract_sales_from_pdf(filepath)

        # Prepare the Excel file for download
        excel_file = os.path.join(app.config['UPLOAD_FOLDER'], 'loreal_sales.xlsx')
        export_to_excel(sales_figure, excel_file)

        # Send the Excel file as a downloadable response
        return send_file(excel_file, as_attachment=True)

# Run the Flask application
if __name__ == "__main__":
    app.run(debug=True)
