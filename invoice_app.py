import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from fpdf import FPDF
from sqlalchemy import Column, Integer, String, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

# ------------------------------
# DATABASE SETUP
# ------------------------------
DB_URL = "sqlite:///customers.db"
engine = create_engine(DB_URL, connect_args={"check_same_thread": False})
Base = declarative_base()
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)


class Customer(Base):
    __tablename__ = "customers"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    address = Column(String)
    email = Column(String)
    vat = Column(String)

    def __str__(self):
        return f"{self.name} - {self.email or 'No Email'}"


Base.metadata.create_all(bind=engine)


# ------------------------------
# DATABASE HELPERS
# ------------------------------
def add_customer(name, address, email, vat):
    with SessionLocal() as session:
        customer = Customer(name=name, address=address, email=email, vat=vat)
        session.add(customer)
        session.commit()

def get_customers():
    with SessionLocal() as session:
        return session.query(Customer).order_by(Customer.name).all()


# ------------------------------
# INVOICE GENERATION
# ------------------------------
def create_invoice_pdf(customer, invoice_number, date, total_amount, trip_count):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", "B", 14)

    pdf.cell(200, 10, "INVOICE", ln=True, align="C")
    pdf.set_font("Arial", size=12)
    pdf.ln(10)

    # FROM
    pdf.cell(100, 6, "From:")
    pdf.set_font("Arial", size=10)
    pdf.multi_cell(100, 6, """
Limousine Service Xpress ApS
Industriholmen 82
2650 Hvidovre
CVR: DK45247961
IBAN: LT87 3250 0345 4552 5735
SWIFT: REVOLT21
Email: limoexpresscph@gmail.com
    """, align="L")

    # TO
    pdf.set_xy(120, 30)
    pdf.set_font("Arial", size=12)
    pdf.cell(80, 6, "To:")
    pdf.set_font("Arial", size=10)
    pdf.set_xy(120, 36)
    pdf.multi_cell(80, 6, f"""
{customer.name}
{customer.address or ''}
{customer.vat or ''}
{customer.email or ''}
    ", align="L")

    # Invoice meta
    pdf.set_xy(10, 100)
    pdf.set_font("Arial", size=12)
    pdf.cell(100, 6, f"Invoice #: {invoice_number}", ln=True)
    pdf.cell(100, 6, f"Date: {date}", ln=True)
    pdf.cell(100, 6, "Currency: DKK", ln=True)

    # Description
    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(120, 10, "Description", border=1)
    pdf.cell(20, 10, "Qty", border=1)
    pdf.cell(50, 10, "Total", border=1, ln=True)

    pdf.set_font("Arial", size=12)
    pdf.cell(120, 10, "Transfers in Copenhagen (see specification)", border=1)
    pdf.cell(20, 10, str(trip_count), border=1)
    pdf.cell(50, 10, f"{total_amount:,.2f} DKK", border=1, ln=True)

    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(190, 8, f"Total Amount Due: {total_amount:,.2f} DKK", ln=True)

    return pdf.output(dest="S").encode("latin-1")


# ------------------------------
# STREAMLIT APP
# ------------------------------
st.set_page_config(page_title="Invoice Generator", layout="centered")
st.title("🧾 Invoice Generator")

tab1, tab2 = st.tabs(["Create Invoice", "Customers"])

with tab2:
    st.header("Customer Management")
    with st.form("add_cust"):
        name = st.text_input("Name")
        address = st.text_area("Address")
        email = st.text_input("Email")
        vat = st.text_input("VAT / Reg No.")
        submit = st.form_submit_button("Add Customer")
        if submit:
            add_customer(name, address, email, vat)
            st.success("Customer added.")

    st.subheader("Customer List")
    st.dataframe(pd.DataFrame([{"Name": c.name, "Email": c.email, "VAT": c.vat} for c in get_customers()]))

with tab1:
    st.header("Generate a New Invoice")

    customer = st.selectbox("Select Customer", options=get_customers())
    invoice_number = st.text_input("Invoice Number", value="")
    invoice_date = st.date_input("Invoice Date", value=datetime.date.today())
    amount = st.number_input("Total Amount (DKK)", min_value=0.0, format="%.2f")
    trips = st.number_input("Trip Count", min_value=1)

    if st.button("Generate PDF"):
        pdf = create_invoice_pdf(customer, invoice_number, invoice_date, amount, trips)
        st.download_button("📥 Download Invoice", data=pdf, file_name=f"Invoice_{invoice_number}.pdf", mime="application/pdf")
