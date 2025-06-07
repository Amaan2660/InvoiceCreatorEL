import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from fpdf import FPDF
from sqlalchemy import Column, Integer, String, Boolean, create_engine
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
    contact = Column(String)
    extra_email = Column(String)
    is_company = Column(Boolean, default=True)

    def __str__(self):
        return f"{self.name} ({'Company' if self.is_company else 'Private'})"


Base.metadata.create_all(bind=engine)


# ------------------------------
# DATABASE HELPERS
# ------------------------------
def add_customer(name, address, email, vat, contact, extra_email, is_company):
    with SessionLocal() as session:
        customer = Customer(
            name=name,
            address=address,
            email=email,
            vat=vat,
            contact=contact,
            extra_email=extra_email,
            is_company=is_company
        )
        session.add(customer)
        session.commit()

def get_customers():
    with SessionLocal() as session:
        return session.query(Customer).order_by(Customer.name).all()

def delete_customer(customer_id):
    with SessionLocal() as session:
        customer = session.query(Customer).get(customer_id)
        if customer:
            session.delete(customer)
            session.commit()

def update_customer(customer_id, **fields):
    with SessionLocal() as session:
        customer = session.query(Customer).get(customer_id)
        if customer:
            for key, value in fields.items():
                setattr(customer, key, value)
            session.commit()


# ------------------------------
# INVOICE GENERATION
# ------------------------------
def create_invoice_pdf(customer, invoice_number, date, total_amount, trip_count):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", "B", 14)

    # Logo
    pdf.image("image.png", x=80, y=10, w=50)
    pdf.ln(35)

    pdf.cell(200, 10, "INVOICE", ln=True, align="C")
    pdf.set_font("Arial", size=12)
    pdf.ln(5)

    # FROM
    pdf.set_font("Arial", "B", 11)
    pdf.cell(95, 6, "From:", ln=0)
    pdf.cell(95, 6, "To:", ln=1)

    pdf.set_font("Arial", size=10)
    pdf.cell(95, 6, "Limousine Service Xpress ApS", ln=0)
    pdf.cell(95, 6, customer.name, ln=1)

    pdf.cell(95, 6, "Industriholmen 82", ln=0)
    pdf.cell(95, 6, customer.address or "", ln=1)

    pdf.cell(95, 6, "2650 Hvidovre", ln=0)
    vat_line = customer.vat if customer.is_company else ""
    pdf.cell(95, 6, vat_line, ln=1)

    pdf.cell(95, 6, "CVR: DK45247961", ln=0)
    pdf.cell(95, 6, customer.contact or "", ln=1)

    pdf.cell(95, 6, "IBAN: LT87 3250 0345 4552 5735", ln=0)
    pdf.cell(95, 6, customer.email or "", ln=1)

    pdf.cell(95, 6, "SWIFT: REVOLT21", ln=0)
    pdf.cell(95, 6, customer.extra_email or "", ln=1)

    pdf.cell(95, 6, "Email: limoexpresscph@gmail.com", ln=1)
    pdf.ln(5)

    # Invoice meta
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
st.set_page_config(page_title="InvoiceCreatorEL", layout="centered")
st.title("🚐 InvoiceCreatorEL")

tab1, tab2 = st.tabs(["Create Invoice", "Customers"])

with tab2:
    st.header("Customer Management")
    with st.form("add_cust"):
        name = st.text_input("Name")
        is_company = st.checkbox("Is Company?", value=True)
        address = st.text_area("Address")
        email = st.text_input("Email")
        vat = st.text_input("VAT / Reg No.") if is_company else ""
        contact = st.text_input("Contact Person (optional)")
        extra_email = st.text_input("Additional Email (optional)")
        submit = st.form_submit_button("Add Customer")
        if submit:
            if not name.strip():
                st.error("Customer name is required.")
            else:
                add_customer(name, address, email, vat, contact, extra_email, is_company)
                st.success("Customer added.")

    st.subheader("Customer List")
    customers = get_customers()
    if customers:
        df = pd.DataFrame([{
            "ID": c.id, "Name": c.name, "Email": c.email,
            "VAT": c.vat, "Company": c.is_company
        } for c in customers])
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        if st.button("Delete Selected Customer"):
            ids = df["ID"].tolist()
            if ids:
                delete_customer(ids[-1])
                st.success("Deleted last customer in list (demo delete). Reload to see changes.")
    else:
        st.info("No customers found.")

with tab1:
    st.header("Generate a New Invoice")

    customers = get_customers()
    customer = st.selectbox("Select Customer", options=customers if customers else [])
    invoice_number = st.text_input("Invoice Number", value="")
    invoice_date = st.date_input("Invoice Date", value=datetime.date.today())
    amount = st.number_input("Total Amount (DKK)", min_value=0.0, format="%.2f")
    trips = st.number_input("Trip Count", min_value=1)

    if st.button("Generate PDF"):
        errors = []
        if not customer:
            errors.append("Customer must be selected.")
        if not invoice_number.strip():
            errors.append("Invoice number is required.")
        if amount <= 0:
            errors.append("Total amount must be greater than zero.")

        if errors:
            for err in errors:
                st.error(err)
        else:
            pdf = create_invoice_pdf(customer, invoice_number, invoice_date, amount, trips)
            st.download_button("📥 Download Invoice", data=pdf, file_name=f"Invoice_{invoice_number}.pdf", mime="application/pdf")
