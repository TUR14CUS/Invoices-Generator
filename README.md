## Invoice PDF Generator from Excel Files using Python

This Python script utilizes the `fpdf` library to generate PDF invoices from Excel files located in the "invoices" directory. Each Excel file represents an invoice and is converted into a PDF document with detailed information about the products, total prices, and a summary.

### Requirements
- Python 3.x
- Install the `fpdf` and `pandas` libraries using:
  ```bash
  pip install fpdf pandas
  ```

### Usage
1. Place your Excel invoice files (with a `.xlsx` extension) in the "invoices" directory.
2. Run the script using:
    ```bash
    python script_name.py
    ```
3. The script will convert each Excel invoice into a corresponding PDF and save it in the "pdfs" directory.

### Script Explanation

1. **Import Libraries**
   - The script imports necessary libraries: `fpdf`, `pandas`, `glob`, and `pathlib`.

2. **Retrieve Filepaths**
   - The script retrieves the filepaths of all Excel files in the "invoices" directory using `glob`.

3. **Generate PDF for Each Invoice**
   - For each Excel file, a new PDF is generated.
   - Extract invoice number and date from the filename.
   - Read the Excel file into a pandas DataFrame.
   - Iterate over DataFrame rows and populate the PDF with product details.
   - Calculate and display the total sum of the invoice.
   - Save the PDF in the "pdfs" directory with the same filename.

4. **Customization**
   - You can customize the script by adjusting the column names used in the PDF generation (e.g., 'product_id', 'product_name', 'amount_purchased', 'price_per_unit', 'total_price').

5. **Additional Customization Options**
   - You can add a company name and logo to the invoice by modifying the corresponding sections in the script.

### Note
- Ensure that the "invoices" directory contains valid Excel files with the expected structure.
- Check the "pdfs" directory for the generated PDF invoices.
- Customize the script to fit your specific invoice format, company details, and branding.

Feel free to modify the script to suit your specific invoicing requirements.
