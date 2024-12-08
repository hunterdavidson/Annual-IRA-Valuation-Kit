import pandas as pd
from jinja2 import Template
from weasyprint import HTML
from datetime import datetime
import shutil
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

# Configuration Parameters
EXCEL_FILE = "input_data.xlsx"
OUTPUT_DIR = "output_pdfs"

try:
    shutil.rmtree(OUTPUT_DIR)

except:
    print("No output directory to remove")

os.makedirs(OUTPUT_DIR, exist_ok=True)
SIGNATURE_IMAGE_PATH = "signature.png"

signature_path = f"file:\\{os.path.abspath(SIGNATURE_IMAGE_PATH)}"

# HTML Template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: Arial, sans-serif;
            line-height: 1.2;
            width: 100%;
            margin: 0;
            padding: 0;
            }
        h1 {
            font-size: 20px;
        }
        p {
            margin-bottom: 40px;
            font-size: 14px;
            line-height: 1.2;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            word-wrap: break-word;
            overflow-wrap: break-word;
            page-break-inside: auto;
        }
        th, td {
            padding: 4px;
            text-align: left;
            line-height: 1.2;
            border: none;
            font-size: 14px;
        }
        th {
            border-bottom: 1px solid black;
            vertical-align: bottom;
        }
        th:nth-child(1), td:nth-child(1) {
            width: 36%;
        }
        th:nth-child(2), td:nth-child(2),
        th:nth-child(3), td:nth-child(3),
        th:nth-child(4), td:nth-child(4),
        th:nth-child(5), td:nth-child(5) {
            width: 16%;
        }
        .th-center {
            text-align: center;
        }
        .td-center {
            text-align: center;
            vertical-align: top;
        }
        .td-right {
            text-align: right;
            vertical-align: top;
            padding-right: 20px;
        }
        thead {
            display: table-header-group;
        }
        tfoot {
            display: table-footer-group;
        }
        tbody {
            display: table-row-group;
        }
        .closing {
            margin-top: 50px;
            page-break-inside: avoid;
        }
        .signature {
            margin: 0;
            vertical-align: middle;
            height: 3.6em;
        }
        @page {
            size: Letter;
            margin: 35mm 20mm 35mm 20mm;
            @top-center {
                content: element(header);
            }
            @bottom-center {
                content: element(footer);
            }
        }
        #header h1 {
            margin: 0;
        }

        #header p {
            margin: 0;
            padding: 0;
        }
        #header {
            position: running(header);
            text-align: center;
            font-size: 14px;
            overflow: hidden;
        }
        #footer {
            position: running(footer);
            text-align: center;
            font-size: 14px;
            overflow: hidden;
        }
        .page-number::after {
            content: counter(page);
        }

        .total-pages::after {
            content: counter(pages);
        }
    </style>
</head>   
<body>
    <div id="header">
        <h1>{{ fund }}</h1>
        <p>{{ company_address }}</p>
    </div>
    <div id="footer">
        <p>P: {{ phone }} | F: {{ fax }} | E: {{ email }}<br>
        Page <span class="page-number"></span> of <span class="total-pages"></span></p>
    </div>
    <p>{{ date }}</p>
    <p><b>{{ custodian }}</b><br>
    {% if care_of != "" %}
        {{ care_of }}<br>
    {% endif %}
    {{ address1 }}<br>
    {{ address2 }}</p>
    <p>RE: {{ fund }}</p>
    <p>Dear Sir/Madame:</p>
    <p>Regarding the above-referenced asset investment, please find the following:</p>
    <table>
        <thead>
            <tr>
                <th>IRA Account Name and Number</th>
                <th class="th-center">Investment<br>Date</th>
                <th class="th-center">Investment<br>Amount</th>
                <th class="th-center">Valuation<br>Date</th>
                <th class="th-center">Valuation<br>Amount</th>
            </tr>
        </thead>
        <tbody>
            {% for row in data %}
            <tr>
                <td>{{ row['IRA Account Name and Number'] }}</td>
                <td class="td-center">{{ row['Investment Date'] }}</td>
                <td class="td-right">{{ row['Investment Amount'] }}</td>
                <td class="td-center">{{ row['Valuation Date'] }}</td>
                <td class="td-right">{{ row['Valuation Amount'] }}</td>
            </tr>
            {% endfor %}
        </tbody>
        <tfoot>
            <tr>
                <td colspan="5" style="border-top: 1px solid black;"> </td>
            </tr>
        </tfoot>
    </table>
    <div class="closing">
        <p>Please let me know if you require any other information.</p>
        <p>Sincerely,<br>
        <img src="{{ signature }}" alt="Signature" class="signature"><br>
        {{ printed_name }}<br>{{ title }}</p>
    </div>
</body>
</html>
"""

def read_excel_data(excel_path):
    """Reads the Excel file and returns a cleaned pandas DataFrame."""
    try:
        # Read the Excel file
        df = pd.read_excel(excel_path, sheet_name="Pull Data")

        # Replace "empty row" with blanks
        df.replace("Empty Row", "", inplace=True)

        # Drop rows that are entirely empty
        df.dropna(how='all', inplace=True)

        # Identify problematic rows in 'Investment Date' and 'Valuation Date'
        invalid_dates = df[~df['Investment Date'].apply(lambda x: isinstance(x, pd.Timestamp) or pd.api.types.is_string_dtype(x))]

        invalid_dates = df[~df['Valuation Date'].apply(lambda x: isinstance(x, pd.Timestamp) or pd.api.types.is_string_dtype(x))]

        # Safely convert date columns to datetime, coercing invalid values to NaT
        df['Investment Date'] = pd.to_datetime(df['Investment Date'], errors='coerce')
        df['Valuation Date'] = pd.to_datetime(df['Valuation Date'], errors='coerce')

        # Drop rows where required columns are missing
        df.dropna(subset=['Investment Date', 'Valuation Date'], inplace=True)

        # Format dates
        df['Investment Date'] = df['Investment Date'].dt.strftime('%m/%d/%Y')
        df['Valuation Date'] = df['Valuation Date'].dt.strftime('%m/%d/%Y')

        print(f"Successfully read {len(df)} records from {excel_path}")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

    
def read_excel_aux_data(excel_path):
    """Reads the Excel file and returns a pandas DataFrame in the Master Table Sheet."""
    try:
        df = pd.read_excel(excel_path, sheet_name="Master Table")
        return df
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def sanitize_file_name(file_name, max_length=200):
    """
    Sanitizes file name by removing invalid characters and truncating if necessary.
    """
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        file_name = file_name.replace(char, '')
    # Truncate the file name if it exceeds the max_length
    if len(file_name) > max_length:
        file_name = file_name[:max_length]
    return file_name

def render_html(template_str, group_data, custodian, fund, care_of, address1, address2, printed_name, title, phone, fax, email, company_address):
    """Renders the HTML template with provided data."""
    template = Template(template_str)
    current_date = datetime.now().strftime("%B %d, %Y")
    rendered = template.render(
        date=current_date,
        custodian=custodian,
        fund=fund,
        care_of=care_of,
        address1=address1,
        address2=address2,
        printed_name=printed_name,
        title=title,
        phone=phone,
        fax=fax,
        email=email,
        company_address=company_address,
        data=group_data.to_dict(orient="records"),
        signature=signature_path,
    )
    return rendered


def html_to_pdf(html_content, output_pdf_path):
    """Converts HTML content to a PDF file using WeasyPrint."""
    try:
        HTML(string=html_content).write_pdf(output_pdf_path)
        print(f"PDF successfully created at {output_pdf_path}")
    except Exception as e:
        print(f"Error generating PDF: {e}")

def create_notification_overlay(output_path, num_pages):
    """
    Creates a PDF overlay with a notification for all pages except the last page.
    """
    c = canvas.Canvas(output_path, pagesize=letter)
    width, height = letter

    for page_number in range(1, num_pages):
        # Add the notification text to all pages except the last
        c.drawString(30, 15, "Signature is on the last page.")
        c.showPage()  # Move to the next page

    # Close the canvas
    c.save()
    print(f"Notification overlay created at {output_path}")

def merge_pdfs(original_pdf_path, overlay_pdf_path, final_pdf_path):
    """
    Merges the original PDF with the notification overlay PDF using PyPDF2.
    """
    original = PdfReader(original_pdf_path)
    overlay = PdfReader(overlay_pdf_path)
    writer = PdfWriter()

    for i, page in enumerate(original.pages):
        # Only add overlay to pages before the last
        if i < len(original.pages) - 1:
            page.merge_page(overlay.pages[i])
        writer.add_page(page)

    # Save the merged PDF
    with open(final_pdf_path, "wb") as out_file:
        writer.write(out_file)

    print(f"Final PDF with notifications saved at {final_pdf_path}")


def main():
    data = read_excel_data(EXCEL_FILE)
    aux_data = read_excel_aux_data(EXCEL_FILE)

    if data.empty:
        print("No data to process. Exiting.")
        return

    # Replace NaN or missing values in 'c/o' with a placeholder string
    data['c/o'] = data['c/o'].fillna("")

    printed_name, title, company_address, phone, fax, email = aux_data.iloc[0, 0:6]

    # Group by 'c/o', 'Custodian', and 'Fund'
    grouped = data.groupby(['c/o', 'Custodian', 'Fund'])

    for (care_of, custodian, fund), group in grouped:
        address1 = group.iloc[0]['Street Name']
        address2 = group.iloc[0]['City, State, Zip']

        # Construct sanitized output file names
        sanitized_care_of = sanitize_file_name(care_of.replace(' ', '_'))
        sanitized_fund = sanitize_file_name(fund.replace(' ', '_'))
        sanitized_custodian = sanitize_file_name(custodian.replace(' ', '_'))

        intermediate_pdf = os.path.join(
            OUTPUT_DIR,
            f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}_intermediate.pdf"
        )
        final_pdf = os.path.join(
            OUTPUT_DIR,
            f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}.pdf"
        )
        overlay_pdf = os.path.join(
            OUTPUT_DIR,
            f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}_overlay.pdf"
        )

        # Render and generate the intermediate PDF
        rendered_html = render_html(html_template, group, custodian, fund, care_of, address1, address2, printed_name, title, phone, fax, email, company_address)
        html_to_pdf(rendered_html, intermediate_pdf)

        # Get the number of pages in the generated PDF
        reader = PdfReader(intermediate_pdf)
        num_pages = len(reader.pages)

        if num_pages > 1:
            # Create a notification overlay if there are multiple pages
            create_notification_overlay(overlay_pdf, num_pages)
            # Merge the overlay with the original PDF
            merge_pdfs(intermediate_pdf, overlay_pdf, final_pdf)
            # Clean up temporary files
            os.remove(intermediate_pdf)
            os.remove(overlay_pdf)
        else:
            # If there's only one page, just rename the intermediate PDF to final
            os.rename(intermediate_pdf, final_pdf)

    print("All PDFs have been successfully created.")


if __name__ == "__main__":
    main()
