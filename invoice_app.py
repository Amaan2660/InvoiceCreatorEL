# Full updated Streamlit app code for InvoiceCreatorEL

import datetime
import pandas as pd
import streamlit as st
from fpdf import FPDF
from sqlalchemy import Column, String, Integer, Boolean, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
from io import BytesIO

# DATABASE SETUP
DB_URL = st.secrets["SUPABASE_DB_URL"]
engine = create_engine(DB_URL)
Base = declarative_base()
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)

# MODELS
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

# DB FUNCTIONS
def add_customer(**kwargs):
    with SessionLocal() as session:
        customer = Customer(**kwargs)
        session.add(customer)
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

# PDF GENERATION

def generate_invoice(receiver, invoice_number, items, currency, purpose):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "INVOICE", ln=True)

    today = datetime.date.today().strftime("%Y-%m-%d")
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 8, f"Invoice #: {invoice_number}", ln=True)
    pdf.cell(0, 8, f"Date: {today}", ln=True)
    pdf.ln(4)

    if purpose:
        pdf.multi_cell(0, 6, f"{purpose}")
        pdf.ln(4)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 8, "Bill To:", ln=True)
    pdf.set_font("Helvetica", "", 12)
    pdf.multi_cell(0, 6, f"{receiver.name}\n{receiver.address}")
    if receiver.contact:
        pdf.cell(0, 6, f"Contact: {receiver.contact}", ln=True)
    if receiver.vat and receiver.is_company:
        pdf.cell(0, 6, f"VAT No: {receiver.vat}", ln=True)
    pdf.cell(0, 6, f"Email: {receiver.email}", ln=True)
    pdf.ln(6)

    total = 0
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(100, 8, "Description", border=1)
    pdf.cell(40, 8, "Amount", border=1, ln=True)
    pdf.set_font("Helvetica", "", 12)
    for _, row in items.iterrows():
        desc = row.get("Description", "")
        amt = float(row.get("Amount", 0))
        total += amt
        pdf.cell(100, 8, str(desc), border=1)
        pdf.cell(40, 8, f"{currency} {amt:.2f}", border=1, ln=True)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(100, 8, "Total", border=1)
    pdf.cell(40, 8, f"{currency} {total:.2f}", border=1, ln=True)
    return pdf.output(dest="S").encode("latin-1")

# STREAMLIT UI
st.set_page_config("InvoiceCreatorEL", layout="centered")
st.title("📄 Invoice Creator EL")
tab1, tab2 = st.tabs(["🧾 Create Invoice", "👥 Manage Customers"])

# --- CUSTOMER MANAGEMENT ---
with tab2:
    st.subheader("Create New Customer")
    with st.form("add_customer"):
        name = st.text_input("Name")
        is_company = st.radio("Type", ["Company", "Individual"]) == "Company"
        address = st.text_area("Address")
        email = st.text_input("Email")
        contact = st.text_input("Contact Person (optional)")
        vat = st.text_input("VAT No.", disabled=not is_company)
        submitted = st.form_submit_button("Add Customer")
        if submitted and name and email:
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

# --- CREATE INVOICE ---
with tab1:
    st.subheader("Create Invoice")
    customers = get_customers()
    receiver = st.selectbox("Select Customer", customers, format_func=lambda x: x.name if x else "")
    invoice_number = st.text_input("Invoice Number")
    currency = st.selectbox("Currency", ["DKK", "EUR", "USD", "GBP"])
    invoice_purpose = st.text_input("Invoice Description (e.g. For services in June)")
    manual_total = st.number_input("Manual Total Amount (optional)", min_value=0.0, step=100.0)
    manual_bookings = st.number_input("Manual Number of Bookings (optional)", min_value=0)
    uploaded = st.file_uploader("Upload CSV with Description + Amount columns", type=["csv"])
    if uploaded:
        df = pd.read_csv(uploaded)
        st.write("Edit invoice rows below if needed:")
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    if st.button("Generate Invoice"):
        if not receiver or not invoice_number or uploaded is None:
            st.error("Missing required information.")
        else:
            total_amount = manual_total if manual_total > 0 else df["Amount"].sum()
            if manual_total > 0:
                df.loc[len(df)] = ["Manual entry", manual_total]

            pdf_bytes = generate_invoice(receiver, invoice_number, df, currency, invoice_purpose)
            # Specification Excel
            export_df = df.copy()
            export_df = export_df.rename(columns={
                "Date": "Date", "From": "From", "To": "To",
                "Customer Reference": "Customer Reference", "Amount": "Price"
            })
            cleaned = export_df[["Date", "From", "To", "Customer Reference", "Price"]].copy()
            buffer = BytesIO()
            cleaned.to_excel(buffer, index=False)

            st.download_button(
                label="⬇️ Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"Invoice {invoice_number} for {receiver.name}.pdf",
                mime="application/pdf"
            )
            st.download_button(
                label="⬇️ Download Specification XLSX",
                data=buffer.getvalue(),
                file_name=f"SERVICE SPECIFICATION FOR INVOICE {invoice_number}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
