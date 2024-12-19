import pandas as pd
from jinja2 import Template
from weasyprint import HTML
from datetime import datetime
import shutil
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox

# Determine the directory of the executable (the .exe file)
BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))

EXCEL_FILE = os.path.join(BASE_DIR, "input_data.xlsx")
SIGNATURE_IMAGE_PATH = os.path.join(BASE_DIR, "signature.png")
OUTPUT_DIR = os.path.join(BASE_DIR, "output_pdfs")
SINGLE_OUTPUT_DIR = os.path.join(BASE_DIR, "single_output")

print("BASE_DIR:", BASE_DIR)
print("EXCEL_FILE:", EXCEL_FILE)
print("SIGNATURE_IMAGE_PATH:", SIGNATURE_IMAGE_PATH)

# Verify required files exist
if not os.path.exists(EXCEL_FILE):
    print(f"Error: Excel file not found at {EXCEL_FILE}")
    sys.exit(1)

if not os.path.exists(SIGNATURE_IMAGE_PATH):
    print(f"Error: Signature image not found at {SIGNATURE_IMAGE_PATH}")
    sys.exit(1)

# Clean up old directories if they exist
try:
    shutil.rmtree(OUTPUT_DIR)
except FileNotFoundError:
    pass

# Create output directories
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SINGLE_OUTPUT_DIR, exist_ok=True)

try:
    shutil.rmtree(OUTPUT_DIR)
except FileNotFoundError:
    print("No output directory to remove")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(SINGLE_OUTPUT_DIR, exist_ok=True)

signature_path = f"file://{os.path.abspath(SIGNATURE_IMAGE_PATH)}"

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
    try:
        df = pd.read_excel(excel_path, sheet_name="Pull Data")
        df.replace("Empty Row", "", inplace=True)
        df.dropna(how='all', inplace=True)

        df['Investment Date'] = pd.to_datetime(df['Investment Date'], errors='coerce')
        df['Valuation Date'] = pd.to_datetime(df['Valuation Date'], errors='coerce')
        df.dropna(subset=['Investment Date', 'Valuation Date'], inplace=True)

        df['Investment Date'] = df['Investment Date'].dt.strftime('%m/%d/%Y')
        df['Valuation Date'] = df['Valuation Date'].dt.strftime('%m/%d/%Y')
        df['c/o'] = df['c/o'].fillna("")

        # Format Investment Amount and Valuation Amount as currency
        df['Investment Amount'] = df['Investment Amount'].apply(lambda x: '${:,.2f}'.format(x))
        df['Valuation Amount'] = df['Valuation Amount'].apply(lambda x: '${:,.2f}'.format(x))

        print(f"Successfully read {len(df)} records from {excel_path}")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def read_excel_aux_data(excel_path):
    try:
        df = pd.read_excel(excel_path, sheet_name="Master Table")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

def sanitize_file_name(file_name, max_length=200):
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        file_name = file_name.replace(char, '')
    if len(file_name) > max_length:
        file_name = file_name[:max_length]
    return file_name

def render_html(template_str, group_data, custodian, fund, care_of, address1, address2, printed_name, title, phone, fax, email, company_address):
    template = Template(template_str)
    current_date = datetime.now().strftime("%B %d, %Y")
    
    # Just use a relative path for the signature image and set base_url later
    signature_filename = "signature.png"
    
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
        signature=signature_filename,  # Use relative path
    )
    return rendered


def html_to_pdf(html_content, output_pdf_path):
    try:
        # Provide the base_url so that the relative path to signature can be resolved
        HTML(string=html_content, base_url=BASE_DIR).write_pdf(output_pdf_path)
        print(f"PDF successfully created at {output_pdf_path}")
    except Exception as e:
        print(f"Error generating PDF: {e}")

def create_notification_overlay(output_path, num_pages):
    c = canvas.Canvas(output_path, pagesize=letter)
    for page_number in range(1, num_pages):
        c.drawString(30, 15, "Signature is on the last page.")
        c.showPage()

    c.save()
    print(f"Notification overlay created at {output_path}")

def merge_pdfs(original_pdf_path, overlay_pdf_path, final_pdf_path):
    original = PdfReader(original_pdf_path)
    overlay = PdfReader(overlay_pdf_path)
    writer = PdfWriter()

    for i, page in enumerate(original.pages):
        if i < len(original.pages) - 1:
            page.merge_page(overlay.pages[i])
        writer.add_page(page)

    with open(final_pdf_path, "wb") as out_file:
        writer.write(out_file)

    print(f"Final PDF with notifications saved at {final_pdf_path}")

def show_progress(total):
    """Create a small progress window for PDF generation with consistent theme colors."""
    progress_win = tk.Toplevel()
    progress_win.title("Generating PDFs...")
    progress_win.geometry("300x120")
    progress_win.resizable(False, False)
    progress_win.configure(bg="#FCF7F8")

    style = ttk.Style(progress_win)
    style.theme_use("clam")
    style.configure("TLabel", background="#FCF7F8", foreground="black", font=("Arial", 12))
    style.configure("TProgressbar", background="#A9A9A9", troughcolor="#FCF7F8", thickness=10)

    ttk.Label(progress_win, text="Generating PDFs, please wait...").pack(pady=10)
    pbar = ttk.Progressbar(progress_win, orient='horizontal', length=200, mode='determinate', style="TProgressbar")
    pbar.pack(pady=5)
    count_label = ttk.Label(progress_win, text="0 / {0}".format(total))
    count_label.pack()

    progress_win.update()
    return progress_win, pbar, count_label

def generate_all_pdfs(data, aux_data):
    if data.empty:
        print("No data to process. Exiting.")
        return

    printed_name, title, company_address, phone, fax, email = aux_data.iloc[0, 0:6]
    grouped = data.groupby(['c/o', 'Custodian', 'Fund'])

    total_count = len(grouped)
    progress_win, pbar, count_label = show_progress(total_count)

    pbar['maximum'] = total_count
    current_count = 0

    for (care_of, custodian, fund), group in grouped:
        address1 = group.iloc[0]['Street Name']
        address2 = group.iloc[0]['City, State, Zip']

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

        rendered_html = render_html(html_template, group, custodian, fund, care_of, address1, address2, printed_name, title, phone, fax, email, company_address)
        html_to_pdf(rendered_html, intermediate_pdf)

        reader = PdfReader(intermediate_pdf)
        num_pages = len(reader.pages)

        if num_pages > 1:
            create_notification_overlay(overlay_pdf, num_pages)
            merge_pdfs(intermediate_pdf, overlay_pdf, final_pdf)
            os.remove(intermediate_pdf)
            os.remove(overlay_pdf)
        else:
            os.rename(intermediate_pdf, final_pdf)

        current_count += 1
        pbar['value'] = current_count
        count_label.config(text=f"{current_count} / {total_count}")
        progress_win.update()

    progress_win.destroy()
    print("All PDFs have been successfully created.")

def generate_single_pdf(data, aux_data, custodian, fund, care_of):
    printed_name, title, company_address, phone, fax, email = aux_data.iloc[0, 0:6]

    filtered = data[(data['Custodian'] == custodian) & (data['Fund'] == fund) & (data['c/o'] == care_of)]

    if filtered.empty:
        print("No matching data found. Cannot generate single PDF.")
        return

    address1 = filtered.iloc[0]['Street Name']
    address2 = filtered.iloc[0]['City, State, Zip']

    sanitized_care_of = sanitize_file_name(care_of.replace(' ', '_'))
    sanitized_fund = sanitize_file_name(fund.replace(' ', '_'))
    sanitized_custodian = sanitize_file_name(custodian.replace(' ', '_'))

    intermediate_pdf = os.path.join(
        SINGLE_OUTPUT_DIR,
        f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}_intermediate.pdf"
    )
    final_pdf = os.path.join(
        SINGLE_OUTPUT_DIR,
        f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}.pdf"
    )
    overlay_pdf = os.path.join(
        SINGLE_OUTPUT_DIR,
        f"{sanitized_fund}_{sanitized_custodian}_{sanitized_care_of}_overlay.pdf"
    )

    rendered_html = render_html(html_template, filtered, custodian, fund, care_of, address1, address2, printed_name, title, phone, fax, email, company_address)
    html_to_pdf(rendered_html, intermediate_pdf)

    reader = PdfReader(intermediate_pdf)
    num_pages = len(reader.pages)

    if num_pages > 1:
        create_notification_overlay(overlay_pdf, num_pages)
        merge_pdfs(intermediate_pdf, overlay_pdf, final_pdf)
        os.remove(intermediate_pdf)
        os.remove(overlay_pdf)
    else:
        os.rename(intermediate_pdf, final_pdf)

    print("Single PDF has been successfully created.")

def run_gui(data, aux_data):
    grouped = data.groupby(['Custodian', 'Fund', 'c/o']).size().reset_index().drop(columns=0)
    custodians = sorted(grouped['Custodian'].unique().tolist())

    def enable_combobox(cb, enable=True):
        style_name = "Large.TCombobox" if enable else "LargeDisabled.TCombobox"
        cb.configure(state="readonly" if enable else "disabled", style=style_name)

    def on_generate_all():
        generate_all_pdfs(data, aux_data)
        messagebox.showinfo("Success", "All PDFs generated successfully.")

    def on_custodian_select(event):
        selected_cust = custodian_cb.get()
        if selected_cust == "Select a Custodian":
            return
        fund_options = sorted(grouped[grouped['Custodian'] == selected_cust]['Fund'].unique().tolist())
        fund_cb['values'] = fund_options
        fund_cb.set("Select a Fund")

        co_cb.set("Select c/o")
        co_cb['values'] = []
        enable_combobox(co_cb, enable=False)

    def on_fund_select(event):
        selected_cust = custodian_cb.get()
        selected_fund = fund_cb.get()
        if selected_fund == "Select a Fund":
            return
        co_options = grouped[(grouped['Custodian'] == selected_cust) & (grouped['Fund'] == selected_fund)]['c/o'].unique().tolist()
        co_options = sorted(co_options)

        if len(co_options) == 0:
            co_cb.set("No c/o available")
            co_cb['values'] = []
            enable_combobox(co_cb, enable=False)
        elif len(co_options) == 1 and co_options[0] == "":
            co_cb.set("No c/o required")
            co_cb['values'] = []
            enable_combobox(co_cb, enable=False)
        else:
            if "" in co_options:
                display_options = ["None"] + [c for c in co_options if c != ""]
            else:
                display_options = co_options

            co_cb['values'] = display_options
            co_cb.set("Select c/o")
            enable_combobox(co_cb, enable=True)

    def on_generate_single():
        selected_cust = custodian_cb.get()
        selected_fund = fund_cb.get()
        selected_co = co_cb.get()

        if selected_cust == "Select a Custodian" or not selected_cust:
            messagebox.showerror("Error", "Please select a Custodian.")
            return
        if selected_fund == "Select a Fund" or not selected_fund:
            messagebox.showerror("Error", "Please select a Fund.")
            return

        filtered_co_options = grouped[(grouped['Custodian'] == selected_cust) & (grouped['Fund'] == selected_fund)]['c/o'].unique().tolist()

        if len(filtered_co_options) > 1:
            if selected_co == "Select c/o":
                messagebox.showerror("Error", "Please select a c/o option.")
                return
            if selected_co == "None":
                chosen_co = ""
            else:
                chosen_co = selected_co
        else:
            chosen_co = filtered_co_options[0] if len(filtered_co_options) == 1 else ""

        generate_single_pdf(data, aux_data, selected_cust, selected_fund, chosen_co)
        messagebox.showinfo("Success", "Single PDF generated successfully.")

    # Colors and styling
    bg_color = "#FCF7F8"
    label_color = "black"
    button_color = "#A9A9A9"  # Grey color
    font_size = ("Arial", 16)

    root = tk.Tk()
    root.title("PDF Generator")
    root.configure(bg=bg_color)
    root.geometry("600x400")
    root.resizable(False, False)

    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("TButton", background=button_color, foreground=label_color, padding=5, font=font_size)
    style.configure("TLabel", background=bg_color, foreground=label_color, font=font_size, anchor="center")
    style.configure("TFrame", background=bg_color)
    style.configure("Large.TCombobox", fieldbackground="white", background="white", foreground=label_color, font=font_size, padding=(10, 5, 10, 5))
    style.configure("LargeDisabled.TCombobox", fieldbackground="lightgrey", background="lightgrey", foreground="black", font=font_size, padding=(10, 5, 10, 5))

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill="both", expand=True)

    frame.columnconfigure(0, weight=1)
    frame.columnconfigure(1, weight=1)

    lbl_info = ttk.Label(frame, text="PDF Generator")
    lbl_info.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew")

    btn_all = ttk.Button(frame, text="Generate All PDFs", command=on_generate_all)
    btn_all.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")

    line_frame = tk.Frame(frame, bg=label_color, width=550, height=2)
    line_frame.grid(row=2, column=0, columnspan=2, pady=10)

    lbl_custodian = ttk.Label(frame, text="Custodian:")
    lbl_custodian.grid(row=3, column=0, padx=5, pady=10, sticky="e")
    custodian_cb = ttk.Combobox(frame, values=custodians, state="readonly", style="Large.TCombobox", width=40)
    custodian_cb.set("Select a Custodian")
    custodian_cb.grid(row=3, column=1, padx=5, pady=10, sticky="w")

    lbl_fund = ttk.Label(frame, text="Fund:")
    lbl_fund.grid(row=4, column=0, padx=5, pady=10, sticky="e")
    fund_cb = ttk.Combobox(frame, values=[], state="readonly", style="Large.TCombobox", width=40)
    fund_cb.set("Select a Fund")
    fund_cb.grid(row=4, column=1, padx=5, pady=10, sticky="w")

    lbl_co = ttk.Label(frame, text="c/o:")
    lbl_co.grid(row=5, column=0, padx=5, pady=10, sticky="e")
    co_cb = ttk.Combobox(frame, values=[], state="disabled", style="LargeDisabled.TCombobox", width=40)
    co_cb.set("Select c/o")
    co_cb.grid(row=5, column=1, padx=5, pady=10, sticky="w")

    btn_single = ttk.Button(frame, text="Generate Single PDF", command=on_generate_single)
    btn_single.grid(row=6, column=0, columnspan=2, padx=5, pady=20, sticky="ew")

    custodian_cb.bind("<<ComboboxSelected>>", on_custodian_select)
    fund_cb.bind("<<ComboboxSelected>>", on_fund_select)

    root.mainloop()

def main():
    data = read_excel_data(EXCEL_FILE)
    aux_data = read_excel_aux_data(EXCEL_FILE)
    run_gui(data, aux_data)

if __name__ == "__main__":
    main()
