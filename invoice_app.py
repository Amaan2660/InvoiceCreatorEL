# Updated Streamlit Invoice App with Clean PDF Layout and Excel Features

import datetime
import pandas as pd
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
        cust = session.query(Customer).get(id)
        for k, v in updates.items():
            setattr(cust, k, v)
        session.commit()

def delete_customer(id):
    with SessionLocal() as session:
        session.query(Customer).filter(Customer.id == id).delete()
        session.commit()

# ------------------- PDF GENERATION -------------------
def generate_invoice_pdf(receiver, invoice_number, currency, description, total_amount, booking_count, due_date):
    pdf = FPDF()
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 20)
    pdf.set_xy(150, 10)
    pdf.cell(0, 10, f"INVOICE {invoice_number}", ln=True)

    pdf.set_font("Helvetica", size=11)
    pdf.set_xy(10, 30)
    pdf.multi_cell(90, 6, """From:
Limousine Service Xpress ApS
Industriholmen 82
2650 Hvidovre
Denmark
CVR: DK45247961
IBAN: LT87 3250 0345 4552 5735
SWIFT: REVOLT21
Email: limoexpresscph@gmail.com""")

    pdf.set_xy(120, 30)
    to_lines = ["To:", receiver.name]
    if receiver.contact:
        to_lines.append(f"Att: {receiver.contact}")
    if receiver.address:
        to_lines.append(receiver.address)
    to_lines.append(f"Email: {receiver.email}")
    pdf.multi_cell(0, 6, "\n".join(to_lines))

    pdf.set_xy(10, 100)
    today = datetime.date.today().strftime("%Y-%m-%d")
    pdf.cell(0, 6, f"Invoice Date: {today}", ln=True)
    pdf.cell(0, 6, f"Due Date: {due_date}", ln=True)
    pdf.cell(0, 6, f"Currency: {currency}", ln=True)
    pdf.ln(10)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(120, 8, "Description", border=1)
    pdf.cell(30, 8, "Qty", border=1)
    pdf.cell(40, 8, "Total", border=1, ln=True)

    pdf.set_font("Helvetica", size=11)
    pdf.cell(120, 8, description or "Transfers", border=1)
    pdf.cell(30, 8, str(booking_count), border=1)
    pdf.cell(40, 8, f"{total_amount:,.2f} {currency}", border=1, ln=True)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(150, 8, "Subtotal:", border=0)
    pdf.cell(40, 8, f"{total_amount:,.2f} {currency}", ln=True)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(150, 8, "Total Amount Due:", border=0)
    pdf.cell(40, 8, f"{total_amount:,.2f} {currency}", ln=True)

    pdf.ln(10)
    pdf.set_font("Helvetica", style="I", size=10)
    pdf.cell(0, 6, f"Please add invoice number {invoice_number} as reference when making payment.", ln=True)

    return pdf.output(dest="S").encode("latin-1")

# ------------------- STREAMLIT UI -------------------
st.set_page_config("InvoiceCreatorEL", layout="centered")
st.title("\U0001F4C4 Invoice Creator EL")
tab1, tab2 = st.tabs(["\U0001F9FE Create Invoice", "\U0001F465 Manage Customers"])

with tab2:
    st.subheader("Create New Customer")
    with st.form("add_customer"):
        name = st.text_input("Name")
        is_company = st.radio("Type", ["Company", "Individual"]) == "Company"
        address = st.text_area("Address")
        email = st.text_input("Email")
        contact = st.text_input("Contact Person (optional)")
        vat = st.text_input("VAT", disabled=not is_company)
        submitted = st.form_submit_button("Add Customer")
        if submitted:
            if not name or not email:
                st.error("Name and Email are required.")
            else:
                add_customer(name=name, address=address, email=email, contact=contact, vat=vat, is_company=is_company)
                st.success("Customer added.")

    st.subheader("Edit/Delete Customers")
    for cust in get_customers():
        with st.expander(cust.name):
            with st.form(f"edit_{cust.id}"):
                cname = st.text_input("Name", value=cust.name)
                ctype = st.radio("Type", ["Company", "Individual"], index=0 if cust.is_company else 1)
                caddr = st.text_area("Address", value=cust.address)
                cemail = st.text_input("Email", value=cust.email)
                ccontact = st.text_input("Contact", value=cust.contact)
                cvat = st.text_input("VAT", value=cust.vat, disabled=(ctype == "Individual"))
                update = st.form_submit_button("Update")
                delete = st.form_submit_button("Delete", type="primary")
                if update:
                    update_customer(cust.id, {
                        "name": cname, "is_company": (ctype == "Company"),
                        "address": caddr, "email": cemail,
                        "contact": ccontact, "vat": cvat
                    })
                    st.success("Updated.")
                elif delete:
                    delete_customer(cust.id)
                    st.success("Deleted.")

with tab1:
    st.subheader("Create Invoice")
    customers = get_customers()
    receiver = st.selectbox("Select Customer", customers, format_func=lambda x: x.name if x else "")
    invoice_number = st.text_input("Invoice Number")
    currency = st.selectbox("Currency", ["DKK", "EUR", "USD", "GBP"])
    invoice_purpose = st.text_input("Invoice Description (e.g. Transfers in May 2025)")
    due_date = st.date_input("Due Date")

    manual_total = 0.0
    manual_bookings = 0

    auto_mode = st.radio("Invoice Calculation Mode", ["Manual", "Auto from Excel"])

    if auto_mode == "Manual":
        manual_total = st.number_input("Manual Total Amount", min_value=0.0, step=100.0)
        manual_bookings = st.number_input("Manual Number of Bookings", min_value=0)

    uploaded = st.file_uploader("Upload Excel File (Auto Clean)", type=["xlsx"])

    if st.button("Generate Invoice"):
        if not receiver or not invoice_number:
            st.error("Customer and Invoice Number are required.")
        elif not uploaded:
            st.error("Please upload a file.")
        else:
            df = pd.read_excel(uploaded, header=1)
            selected_cols = ["Trip Date", "Passenger", "From", "To", "Customer", "Cust. Ref.", "Base Rate"]
            cleaned_df = df[selected_cols].copy()

            total_amount = manual_total if auto_mode == "Manual" else 0.0
            booking_count = manual_bookings if auto_mode == "Manual" else 0

            if auto_mode == "Auto from Excel":
                booking_count = len(cleaned_df)
                total_amount = booking_count * 395

            cleaned_df.loc[len(cleaned_df.index)] = ["", "", "", "", "", "Total:", cleaned_df["Base Rate"].sum()]

            buffer = BytesIO()
            cleaned_df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            wb = load_workbook(buffer)
            ws = wb.active

            for cell in ws[1]:
                cell.font = Font(bold=True)

            for col in ws.columns:
                max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col[0].column_letter].width = max_length + 2

            final_buffer = BytesIO()
            wb.save(final_buffer)

            st.download_button(
                label="⬇️ Download Specification XLSX",
                data=final_buffer.getvalue(),
                file_name=f"SERVICE SPECIFICATION FOR INVOICE {invoice_number}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            pdf_bytes = generate_invoice_pdf(receiver, invoice_number, currency, invoice_purpose, total_amount, booking_count, due_date.strftime("%Y-%m-%d"))

            st.download_button(
                label="⬇️ Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"Invoice {invoice_number} for {receiver.name}.pdf",
                mime="application/pdf"
            )
