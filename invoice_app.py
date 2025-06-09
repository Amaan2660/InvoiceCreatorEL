import streamlit as st
import pandas as pd
from io import BytesIO
from sqlalchemy import Column, String, Integer, Boolean, create_engine
from sqlalchemy.orm import declarative_base, sessionmaker
import datetime
from fpdf import FPDF

# ---------------- Database Setup ------------------
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

# ---------------- PDF Generator ------------------
def generate_invoice(receiver, invoice_number, items, currency, purpose):
    pdf = FPDF()
    pdf.add_page()
    pdf.image("logo.png", x=10, y=8, w=50)
    pdf.set_xy(120, 8)
    pdf.set_font("Helvetica", size=10)
    pdf.multi_cell(80, 5, """
From:
Limousine Service Xpress ApS
Industriholmen 82
2650 Hvidovre
VAT: 45247961
IBAN: LT87 3250 0345 4552 5735
SWIFT: REVOLT21
Email: limoexpresscph@gmail.com
    """)

    pdf.ln(30)
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "INVOICE", ln=True)

    today = datetime.date.today().strftime("%Y-%m-%d")
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 6, f"Invoice #: {invoice_number}", ln=True)
    pdf.cell(0, 6, f"Date: {today}", ln=True)
    if purpose:
        pdf.multi_cell(0, 6, f"Description: {purpose}")
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 6, "To:", ln=True)
    pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 6, f"{receiver.name}\n{receiver.address}")
    if receiver.contact:
        pdf.cell(0, 6, f"Contact: {receiver.contact}", ln=True)
    if receiver.vat and receiver.is_company:
        pdf.cell(0, 6, f"VAT No: {receiver.vat}", ln=True)
    pdf.cell(0, 6, f"Email: {receiver.email}", ln=True)
    pdf.ln(6)

    total = 0
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(100, 8, "Service", border=1)
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

# ---------------- Streamlit UI ------------------
st.set_page_config("InvoiceCreatorEL", layout="centered")
st.title("📄 Invoice Creator EL")
tab1, tab2 = st.tabs(["🧾 Create Invoice", "👥 Manage Customers"])

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
    manual_total = st.number_input("Manual Total Amount (optional)", min_value=0.0, step=100.0)
    manual_bookings = st.number_input("Manual Number of Bookings (optional)", min_value=0)
    uploaded = st.file_uploader("Upload Excel file with booking data", type=["xlsx"])

    df = pd.DataFrame(columns=["Description", "Amount"])

    if uploaded:
        raw = pd.read_excel(uploaded, header=0)
        expected_cols = ["Trip Date", "Passenger", "From", "To", "Cust. Ref."]
        filtered = raw[[col for col in expected_cols if col in raw.columns]]
        st.write("Preview of cleaned data:")
        st.dataframe(filtered)

        buffer = BytesIO()
        filtered.to_excel(buffer, index=False, engine="openpyxl")
        st.download_button(
            label="⬇️ Download Cleaned Specification XLSX",
            data=buffer.getvalue(),
            file_name=f"SERVICE SPECIFICATION FOR INVOICE {invoice_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    if st.button("Generate Invoice"):
        if not receiver or not invoice_number:
            st.error("Customer and Invoice Number are required.")
        else:
            if manual_total > 0:
                df = pd.DataFrame([{"Description": invoice_purpose, "Amount": manual_total}])
            pdf_bytes = generate_invoice(receiver, invoice_number, df, currency, invoice_purpose)
            st.download_button(
                label="⬇️ Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"Invoice {invoice_number} for {receiver.name}.pdf",
                mime="application/pdf"
            )
