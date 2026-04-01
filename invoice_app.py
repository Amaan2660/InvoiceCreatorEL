# Updated Streamlit Invoice App with Clean PDF Layout and Excel Features

import datetime
import pandas as pd
import base64
from fpdf import FPDF
from io import BytesIO
import streamlit as st
from sqlalchemy import create_engine, Column, String, Integer, Boolean
from sqlalchemy.orm import declarative_base, sessionmaker
from openpyxl import load_workbook
from openpyxl.styles import Font

# ------------------- DATABASE SETUP -------------------
DB_URL = st.secrets["SUPABASE_DB_URL"]
engine = create_engine(DB_URL)
Base = declarative_base()
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)

class Customer(Base):
    __tablename__ = "customers"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    address = Column(String)
    email = Column(String)
    contact = Column(String)
    vat = Column(String)
    is_company = Column(Boolean, default=True)

Base.metadata.create_all(bind=engine)

def add_customer(**kwargs):
    with SessionLocal() as session:
        session.add(Customer(**kwargs))
        session.commit()

def get_customers():
    with SessionLocal() as session:
        return session.query(Customer).order_by(Customer.name).all()

def update_customer(id, updates):
    with SessionLocal() as session:
        cust = session.get(Customer, id)
        if cust:
            for k, v in updates.items():
                setattr(cust, k, v)
            session.commit()

def delete_customer(id):
    with SessionLocal() as session:
        session.query(Customer).filter(Customer.id == id).delete()
        session.commit()

# ------------------- CURRENCY CONVERSION -------------------
def convert_currency(amount_dkk, target_currency):
    rates = {
        "EUR": 7.5,
        "USD": 6.5,
        "GBP": 8.8
    }
    rate = rates.get(target_currency)
    return round(amount_dkk / rate, 2) if rate else amount_dkk

def get_currency_note(currency):
    rates = {
        "EUR": 7.5,
        "USD": 6.5,
        "GBP": 8.8
    }
    return f"{currency} (1 {currency} = {rates[currency]} DKK)" if currency in rates else currency

# ------------------- BANK DETAILS -------------------
def get_bank_details(bank_choice):
    if bank_choice == "Nordea":
        return {
            "bank_name": "Nordea",
            "iban": "DK41 2000 9046 3317 85",
            "swift": "NDEADKKK",
            "reg_no": "2355",
            "account_no": "9046331785"
        }
    return {
        "bank_name": "Revolut",
        "iban": "LT87 3250 0345 4552 5735",
        "swift": "REVOLT21",
        "reg_no": "",
        "account_no": ""
    }

# ------------------- PDF GENERATION -------------------
def generate_invoice_pdf(receiver, invoice_number, currency, description, total_amount, booking_count, due_date, bank_choice):
    pdf = FPDF()
    pdf.add_page()
    bank_details = get_bank_details(bank_choice)

    try:
        pdf.image("logo.png", x=10, y=8, w=40)
    except Exception:
        pass

    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(150, 10)
    pdf.cell(0, 10, f"INVOICE {invoice_number}", ln=True)

    # From block
    pdf.set_font("Helvetica", size=11)
    pdf.set_xy(10, 30)
    pdf.set_font("Helvetica", "B", 11)
    pdf.multi_cell(90, 6, "From:")
    pdf.set_font("Helvetica", "", 11)

    sender_text = f"""Limousine Service Xpress ApS
Industriholmen 82
2650 Hvidovre
Denmark
CVR/VAT: DK45247961
IBAN: {bank_details['iban']}
SWIFT: {bank_details['swift']}"""

    if bank_details["reg_no"]:
        sender_text += f"""
Reg Nr: {bank_details['reg_no']}
Konto Nr: {bank_details['account_no']}"""

    sender_text += """
Email: limoexpresscph@gmail.com"""

    pdf.multi_cell(90, 6, sender_text)

    # To block
    pdf.set_xy(120, 30)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, "To:", ln=True)
    pdf.set_font("Helvetica", "", 11)

    pdf.set_x(120)
    if receiver.name:
        pdf.cell(0, 6, receiver.name, ln=True)
    if receiver.contact:
        pdf.set_x(120)
        pdf.multi_cell(80, 6, f"Att: {receiver.contact}")
    if receiver.address:
        pdf.set_x(120)
        pdf.multi_cell(80, 6, receiver.address)
    if receiver.vat and receiver.is_company:
        pdf.set_x(120)
        pdf.cell(0, 6, f"VAT No: {receiver.vat}", ln=True)
    if receiver.email:
        pdf.set_x(120)
        pdf.multi_cell(80, 6, f"Email: {receiver.email}")

    # Invoice details
    pdf.set_xy(10, 100)
    today = datetime.date.today().strftime("%d/%m/%Y")
    due_date_fmt = due_date.strftime("%d/%m/%Y")
    currency_note = get_currency_note(currency)
    pdf.cell(0, 6, f"Invoice Date: {today}", ln=True)
    pdf.cell(0, 6, f"Due Date: {due_date_fmt}", ln=True)
    pdf.cell(0, 6, f"Currency: {currency_note}", ln=True)
    pdf.ln(10)

    # Invoice table
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(120, 8, "Description", border=1)
    pdf.cell(30, 8, "Qty", border=1)
    pdf.cell(40, 8, "Total", border=1, ln=True)

    pdf.set_font("Helvetica", "", 11)
    pdf.cell(120, 8, description or "Transfers", border=1)
    pdf.cell(30, 8, str(booking_count), border=1)
    pdf.cell(40, 8, f"{total_amount:,.2f} {currency}", border=1, ln=True)

    pdf.ln(8)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(150, 8, "Subtotal:", border=0)
    pdf.cell(40, 8, f"{float(total_amount):,.2f} {currency}", ln=True)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(150, 8, "Total Amount Due:", border=0)
    pdf.cell(40, 8, f"{total_amount:,.2f} {currency}", ln=True)

    pdf.ln(10)
    pdf.set_font("Helvetica", "I", 10)
    pdf.cell(0, 6, f"Please add invoice number {invoice_number} as reference when making payment.", ln=True)

    # Bottom notice for Nordea only
    if bank_choice == "Nordea":
        pdf.set_y(-10)
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(190, 8, "IMPORTANT: PLEASE USE OUR NEW BANK DETAILS", border=1, ln=True, align="C")

        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(
            190,
            6,
            "Please make payment using the Nordea banking information stated on this invoice.",
            border=1,
            align="C"
        )

    return pdf.output(dest="S").encode("latin-1")

# ------------------- PREVIEW HELPERS -------------------
def preview_pdf(bytes_pdf):
    b64 = base64.b64encode(bytes_pdf).decode()
    return f"<iframe src='data:application/pdf;base64,{b64}' width='700' height='900' type='application/pdf'></iframe>"

def preview_excel(df):
    return st.dataframe(df)

# ------------------- STREAMLIT UI -------------------
st.set_page_config(page_title="InvoiceCreatorEL", layout="centered")
st.title("📄 Invoice Creator EL")
tab1, tab2 = st.tabs(["🧾 Create Invoice", "👥 Manage Customers"])

with tab1:
    st.subheader("Create Invoice")
    customers = get_customers()
    receiver = st.selectbox("Select Customer", customers, format_func=lambda x: x.name if x else "")
    invoice_number = st.text_input("Invoice Number")
    currency = st.selectbox("Currency", ["DKK", "EUR", "USD", "GBP"])
    bank_choice = st.selectbox("Bank for payment", ["Nordea", "Revolut"], index=0)
    invoice_purpose = st.text_input("Invoice Description (e.g. Transfers in May 2025)")
    due_date = st.date_input("Due Date")
    mode = st.radio("Select Amount Mode", ["Manual", "Auto from Excel"])

    uploaded = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
    total_amount_dkk = 0.0
    booking_count = 0
    cleaned_df = pd.DataFrame()

    if uploaded:
        try:
            if uploaded.name.endswith(".xls"):
                df = pd.read_excel(uploaded, header=1, engine="xlrd")
            else:
                df = pd.read_excel(uploaded, header=1, engine="openpyxl")

            target_cols = ['Trip Date', 'Passenger', 'From', 'To', 'Customer', 'Cust. Ref.', 'Base Rate']
            missing_cols = [col for col in target_cols if col not in df.columns]

            if missing_cols:
                st.error(f"Missing required columns in Excel file: {', '.join(missing_cols)}")
            else:
                cleaned_df = df[target_cols].copy()
                cleaned_df = cleaned_df.dropna(subset=['Base Rate'])
                cleaned_df['Base Rate'] = cleaned_df['Base Rate'].astype(str).str.replace(',', '', regex=False)
                cleaned_df['Base Rate'] = pd.to_numeric(cleaned_df['Base Rate'], errors='coerce')
                cleaned_df = cleaned_df.dropna(subset=['Base Rate'])

                if not cleaned_df.empty:
                    last_value = cleaned_df['Base Rate'].iloc[-1]
                    sum_except_last = cleaned_df['Base Rate'].iloc[:-1].sum()
                    if len(cleaned_df) > 1 and abs(last_value - sum_except_last) < 1.0:
                        cleaned_df = cleaned_df.iloc[:-1]

                if mode == "Auto from Excel":
                    booking_count = cleaned_df.shape[0]
                    total_amount_dkk = cleaned_df['Base Rate'].sum()
        except Exception as e:
            st.error(f"Could not read Excel file: {e}")

    if mode == "Manual":
        total_amount_dkk = st.number_input("Manual Total Amount", min_value=0.0, step=100.0)
        booking_count = st.number_input("Manual Number of Bookings", min_value=0)

    if st.button("Generate Invoice"):
        if not receiver or not invoice_number:
            st.error("Customer and Invoice Number are required.")
        else:
            if not cleaned_df.empty:
                buffer = BytesIO()
                cleaned_df.to_excel(buffer, index=False, engine="openpyxl")
                buffer.seek(0)

                wb = load_workbook(buffer)
                ws = wb.active
                bold_font = Font(bold=True)

                for cell in ws[1]:
                    cell.font = bold_font

                for col in ws.columns:
                    values = [len(str(cell.value)) for cell in col if cell.value is not None]
                    max_length = max(values) if values else 10
                    ws.column_dimensions[col[0].column_letter].width = max_length + 2

                ws.append(["", "", "", "", "", "Total", cleaned_df['Base Rate'].sum()])

                final_buffer = BytesIO()
                wb.save(final_buffer)

                preview_excel(cleaned_df)
                st.download_button(
                    "⬇️ Download Specification XLSX",
                    data=final_buffer.getvalue(),
                    file_name=f"SERVICE SPECIFICATION FOR INVOICE {invoice_number}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            final_total = convert_currency(total_amount_dkk, currency) if mode == "Auto from Excel" and currency != "DKK" else total_amount_dkk

            pdf_bytes = generate_invoice_pdf(
                receiver=receiver,
                invoice_number=invoice_number,
                currency=currency,
                description=invoice_purpose,
                total_amount=final_total,
                booking_count=booking_count,
                due_date=due_date,
                bank_choice=bank_choice
            )

            st.markdown(preview_pdf(pdf_bytes), unsafe_allow_html=True)
            st.download_button(
                "⬇️ Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"Invoice {invoice_number} for {receiver.name}.pdf",
                mime="application/pdf"
            )

with tab2:
    st.subheader("Manage Customers")

    customers = get_customers()
    if not customers:
        st.info("No customers found. Please add one below.")
    else:
        selected = st.selectbox("Select Customer to Edit/Delete", customers, format_func=lambda c: c.name)

        if selected:
            with st.expander("Edit Customer"):
                name = st.text_input("Name", selected.name)
                email = st.text_input("Email", selected.email)
                address = st.text_input("Address", selected.address)
                contact = st.text_input("Contact", selected.contact)
                vat = st.text_input("VAT Number", selected.vat)
                is_company = st.checkbox("Is Company", selected.is_company)

                if st.button("Update Customer"):
                    update_customer(selected.id, {
                        "name": name,
                        "email": email,
                        "address": address,
                        "contact": contact,
                        "vat": vat,
                        "is_company": is_company
                    })
                    st.success("Customer updated successfully.")

            if st.button("Delete Customer"):
                delete_customer(selected.id)
                st.success("Customer deleted successfully.")

        st.markdown("---")

    with st.expander("Add New Customer"):
        new_name = st.text_input("New Name")
        new_email = st.text_input("New Email")
        new_address = st.text_input("New Address")
        new_contact = st.text_input("New Contact")
        new_vat = st.text_input("New VAT Number")
        new_is_company = st.checkbox("New Is Company", value=True)

        if st.button("Add Customer"):
            if new_name:
                add_customer(
                    name=new_name,
                    email=new_email,
                    address=new_address,
                    contact=new_contact,
                    vat=new_vat,
                    is_company=new_is_company
                )
                st.success("Customer added successfully.")
            else:
                st.warning("Name is required.")
