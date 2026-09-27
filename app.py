import streamlit as st
from datetime import date
from pathlib import Path
import base64
from io import BytesIO

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="LegalEase",
    page_icon="⚖️",
    layout="wide"
)


# ============================================================
# BACKGROUND IMAGE
# ============================================================

def load_background():
    image_path = Path("legal_background.png")

    if image_path.exists():
        encoded = base64.b64encode(
            image_path.read_bytes()
        ).decode()

        return f"data:image/png;base64,{encoded}"

    return ""


background = load_background()


# ============================================================
# CSS ONLY
# ============================================================

st.markdown(
    f"""
    <style>

    .stApp {{
        background-image:
            linear-gradient(
                rgba(255,255,255,0.12),
                rgba(255,255,255,0.12)
            ),
            url("{background}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}

    .main .block-container {{
        max-width: 1400px;
        padding-top: 25px;
        padding-bottom: 50px;
    }}

    /* Header */
    .header-box {{
        background: rgba(255,248,235,0.94);
        border: 1px solid #8b6848;
        border-radius: 15px;
        padding: 22px;
        text-align: center;
        margin-bottom: 22px;
        box-shadow: 0 8px 25px rgba(60,35,15,0.25);
    }}

    /* Main panels */
    .panel {{
        background: rgba(255,250,242,0.96);
        border: 1px solid #a78361;
        border-radius: 14px;
        padding: 28px;
        margin-bottom: 22px;
        box-shadow: 0 8px 25px rgba(60,35,15,0.20);
    }}

    /* Document preview */
    .document-paper {{
        background: #fffdf8;
        border: 1px solid #b8a083;
        padding: 45px;
        box-shadow: 0 8px 25px rgba(50,30,15,0.18);
        min-height: 500px;
    }}

    /* Text */
    .dark-text {{
        color: #21140d;
    }}

    /* Buttons */
    .stButton > button {{
        background-color: #5a321b !important;
        color: white !important;
        border-radius: 8px !important;
        border: 1px solid #3e210f !important;
        font-weight: 700 !important;
        min-height: 45px;
    }}

    .stButton > button:hover {{
        background-color: #754624 !important;
    }}

    /* Inputs */
    .stTextInput input,
    .stTextArea textarea,
    .stDateInput input {{
        background-color: #fffdf9 !important;
        color: #21140d !important;
        border: 1px solid #a78361 !important;
    }}

    /* Labels */
    label {{
        color: #2c190e !important;
        font-weight: 600 !important;
    }}

    div[data-baseweb="select"] {{
        background-color: #fffdf9 !important;
    }}

    div[data-baseweb="select"] * {{
        color: #21140d !important;
    }}

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

with st.container():
    st.markdown(
        "<h1 style='text-align:center; color:#2b180f; "
        "font-family:Georgia,serif;'>⚖️ LegalEase</h1>",
        unsafe_allow_html=True
    )

    st.markdown(
        "<p style='text-align:center; color:#4b2d1c; "
        "font-weight:600; letter-spacing:1px;'>"
        "PROFESSIONAL LEGAL DOCUMENT GENERATOR"
        "</p>",
        unsafe_allow_html=True
    )


# ============================================================
# DOCUMENT TYPE
# ============================================================

st.markdown(
    "<div class='panel'>",
    unsafe_allow_html=True
)

st.subheader("Create Your Legal Document")

st.write(
    "Select the document type and enter the required information. "
    "LegalEase will organize the information into a structured "
    "professional legal document."
)

document_type = st.selectbox(
    "Select Document Type",
    [
        "Employment Agreement",
        "Non-Disclosure Agreement (NDA)",
        "Residential Lease Agreement",
        "Service Agreement",
        "Partnership Agreement"
    ]
)

st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# VARIABLES
# ============================================================

data = {}


# ============================================================
# EMPLOYMENT AGREEMENT
# ============================================================

if document_type == "Employment Agreement":

    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    st.subheader("Employment & Party Details")

    col1, col2 = st.columns(2)

    with col1:
        data["Employer"] = st.text_input(
            "Employer / Company Name"
        )

        data["Employee"] = st.text_input(
            "Employee Full Name"
        )

        data["Job Title"] = st.text_input(
            "Job Title / Position"
        )

        data["Department"] = st.text_input(
            "Department"
        )

    with col2:
        data["Effective Date"] = st.date_input(
            "Effective / Joining Date",
            value=date.today()
        )

        data["Employment Type"] = st.selectbox(
            "Employment Type",
            [
                "Full-Time",
                "Part-Time",
                "Contract",
                "Temporary"
            ]
        )

        data["Salary"] = st.text_input(
            "Salary / Compensation"
        )

        data["Probation"] = st.text_input(
            "Probation Period"
        )

    st.subheader("Employment Conditions")

    data["Working Hours"] = st.text_area(
        "Working Hours"
    )

    data["Leave & Benefits"] = st.text_area(
        "Leave & Benefits"
    )

    data["Confidentiality"] = st.text_area(
        "Confidentiality"
    )

    data["Notice Period"] = st.text_input(
        "Notice Period"
    )

    data["Termination"] = st.text_area(
        "Termination Terms"
    )

    data["Additional Terms"] = st.text_area(
        "Additional Terms & Conditions"
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# NDA
# ============================================================

elif document_type == "Non-Disclosure Agreement (NDA)":

    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    st.subheader("NDA Party Details")

    col1, col2 = st.columns(2)

    with col1:
        data["Disclosing Party"] = st.text_input(
            "Disclosing Party"
        )

        data["Receiving Party"] = st.text_input(
            "Receiving Party"
        )

        data["Purpose"] = st.text_input(
            "Purpose of Agreement"
        )

    with col2:
        data["Effective Date"] = st.date_input(
            "Effective Date",
            value=date.today()
        )

        data["Confidentiality Period"] = st.text_input(
            "Confidentiality Period"
        )

        data["Jurisdiction"] = st.text_input(
            "Governing Jurisdiction"
        )

    st.subheader("Confidentiality Conditions")

    data["Confidential Information"] = st.text_area(
        "Definition of Confidential Information"
    )

    data["Obligations"] = st.text_area(
        "Receiving Party Obligations"
    )

    data["Exceptions"] = st.text_area(
        "Exclusions / Exceptions"
    )

    data["Permitted Disclosure"] = st.text_area(
        "Permitted Disclosure"
    )

    data["Additional Terms"] = st.text_area(
        "Additional NDA Terms"
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# LEASE AGREEMENT
# ============================================================

elif document_type == "Residential Lease Agreement":

    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    st.subheader("Landlord, Tenant & Property")

    col1, col2 = st.columns(2)

    with col1:
        data["Landlord"] = st.text_input(
            "Landlord Full Name"
        )

        data["Tenant"] = st.text_input(
            "Tenant Full Name"
        )

        data["Property Address"] = st.text_area(
            "Property Address"
        )

    with col2:
        data["Lease Start"] = st.date_input(
            "Lease Start Date",
            value=date.today()
        )

        data["Lease End"] = st.date_input(
            "Lease End Date",
            value=date.today()
        )

        data["Property Type"] = st.selectbox(
            "Property Type",
            [
                "Apartment",
                "House",
                "Villa",
                "Studio",
                "Other"
            ]
        )

    st.subheader("Financial Terms")

    col1, col2 = st.columns(2)

    with col1:
        data["Monthly Rent"] = st.text_input(
            "Monthly Rent"
        )

        data["Security Deposit"] = st.text_input(
            "Security Deposit"
        )

    with col2:
        data["Payment Due"] = st.text_input(
            "Rent Payment Due"
        )

        data["Late Fee"] = st.text_input(
            "Late Payment Terms"
        )

    data["Maintenance"] = st.text_area(
        "Maintenance Responsibilities"
    )

    data["Utilities"] = st.text_area(
        "Utilities"
    )

    data["Termination"] = st.text_area(
        "Lease Termination Conditions"
    )

    data["Additional Terms"] = st.text_area(
        "Additional Lease Terms"
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# SERVICE AGREEMENT
# ============================================================

elif document_type == "Service Agreement":

    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    st.subheader("Service Provider & Client")

    col1, col2 = st.columns(2)

    with col1:
        data["Service Provider"] = st.text_input(
            "Service Provider"
        )

        data["Client"] = st.text_input(
            "Client Name"
        )

        data["Service"] = st.text_input(
            "Service / Project Name"
        )

        data["Location"] = st.text_input(
            "Service Location"
        )

    with col2:
        data["Agreement Date"] = st.date_input(
            "Agreement Date",
            value=date.today()
        )

        data["Start Date"] = st.date_input(
            "Service Start Date",
            value=date.today()
        )

        data["Duration"] = st.text_input(
            "Service Duration"
        )

        data["Fee"] = st.text_input(
            "Service Fee"
        )

    st.subheader("Service Conditions")

    data["Scope"] = st.text_area(
        "Scope of Services"
    )

    data["Responsibilities"] = st.text_area(
        "Responsibilities of Parties"
    )

    data["Payment Schedule"] = st.text_input(
        "Payment Schedule"
    )

    data["Intellectual Property"] = st.text_area(
        "Intellectual Property Terms"
    )

    data["Termination"] = st.text_area(
        "Termination Terms"
    )

    data["Additional Terms"] = st.text_area(
        "Additional Service Terms"
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# PARTNERSHIP AGREEMENT
# ============================================================

elif document_type == "Partnership Agreement":

    st.markdown("<div class='panel'>", unsafe_allow_html=True)

    st.subheader("Partnership & Partner Details")

    col1, col2 = st.columns(2)

    with col1:
        data["Partnership Name"] = st.text_input(
            "Partnership / Business Name"
        )

        data["Partner 1"] = st.text_input(
            "Partner 1 Full Name"
        )

        data["Partner 2"] = st.text_input(
            "Partner 2 Full Name"
        )

    with col2:
        data["Agreement Date"] = st.date_input(
            "Agreement Date",
            value=date.today()
        )

        data["Business Activity"] = st.text_input(
            "Nature of Business"
        )

        data["Business Location"] = st.text_input(
            "Principal Business Location"
        )

    st.subheader("Capital & Profit Sharing")

    col1, col2 = st.columns(2)

    with col1:
        data["Capital Partner 1"] = st.text_input(
            "Partner 1 Capital Contribution"
        )

        data["Capital Partner 2"] = st.text_input(
            "Partner 2 Capital Contribution"
        )

    with col2:
        data["Profit Ratio"] = st.text_input(
            "Profit / Loss Sharing Ratio"
        )

        data["Duration"] = st.text_input(
            "Partnership Duration"
        )

    data["Responsibilities"] = st.text_area(
        "Partner Responsibilities"
    )

    data["Decision Making"] = st.text_area(
        "Decision-Making Terms"
    )

    data["Withdrawal"] = st.text_area(
        "Withdrawal / Retirement Terms"
    )

    data["Dissolution"] = st.text_area(
        "Dissolution Terms"
    )

    data["Additional Terms"] = st.text_area(
        "Additional Partnership Terms"
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# DOCUMENT GENERATION FUNCTION
# ============================================================

def create_document(document_type, data):

    lines = []

    lines.append("LEGALEASE")
    lines.append(document_type.upper())
    lines.append("=" * 70)
    lines.append("")

    if document_type == "Employment Agreement":

        lines.extend([
            "1. PARTIES",
            "",
            f"This Employment Agreement is entered into between "
            f"{data.get('Employer', '')} and "
            f"{data.get('Employee', '')}.",
            "",
            "2. POSITION AND EMPLOYMENT",
            "",
            f"Position: {data.get('Job Title', '')}",
            f"Department: {data.get('Department', '')}",
            f"Employment Type: {data.get('Employment Type', '')}",
            f"Effective Date: {data.get('Effective Date', '')}",
            "",
            "3. COMPENSATION",
            "",
            f"Salary / Compensation: {data.get('Salary', '')}",
            "",
            "4. WORKING HOURS",
            "",
            data.get("Working Hours", ""),
            "",
            "5. PROBATION",
            "",
            f"Probation Period: {data.get('Probation', '')}",
            "",
            "6. LEAVE AND BENEFITS",
            "",
            data.get("Leave & Benefits", ""),
            "",
            "7. CONFIDENTIALITY",
            "",
            data.get("Confidentiality", ""),
            "",
            "8. NOTICE PERIOD",
            "",
            f"Notice Period: {data.get('Notice Period', '')}",
            "",
            "9. TERMINATION",
            "",
            data.get("Termination", ""),
            "",
            "10. ADDITIONAL TERMS",
            "",
            data.get("Additional Terms", "")
        ])

    elif document_type == "Non-Disclosure Agreement (NDA)":

        lines.extend([
            "1. PARTIES",
            "",
            f"Disclosing Party: {data.get('Disclosing Party', '')}",
            f"Receiving Party: {data.get('Receiving Party', '')}",
            "",
            "2. PURPOSE",
            "",
            data.get("Purpose", ""),
            "",
            "3. CONFIDENTIAL INFORMATION",
            "",
            data.get("Confidential Information", ""),
            "",
            "4. OBLIGATIONS",
            "",
            data.get("Obligations", ""),
            "",
            "5. EXCLUSIONS AND EXCEPTIONS",
            "",
            data.get("Exceptions", ""),
            "",
            "6. PERMITTED DISCLOSURE",
            "",
            data.get("Permitted Disclosure", ""),
            "",
            "7. CONFIDENTIALITY PERIOD",
            "",
            data.get("Confidentiality Period", ""),
            "",
            "8. GOVERNING JURISDICTION",
            "",
            data.get("Jurisdiction", ""),
            "",
            "9. ADDITIONAL TERMS",
            "",
            data.get("Additional Terms", "")
        ])

    elif document_type == "Residential Lease Agreement":

        lines.extend([
            "1. PARTIES",
            "",
            f"Landlord: {data.get('Landlord', '')}",
            f"Tenant: {data.get('Tenant', '')}",
            "",
            "2. PROPERTY",
            "",
            f"Property Type: {data.get('Property Type', '')}",
            f"Property Address: {data.get('Property Address', '')}",
            "",
            "3. LEASE TERM",
            "",
            f"Lease Start: {data.get('Lease Start', '')}",
            f"Lease End: {data.get('Lease End', '')}",
            "",
            "4. RENT",
            "",
            f"Monthly Rent: {data.get('Monthly Rent', '')}",
            f"Payment Due: {data.get('Payment Due', '')}",
            "",
            "5. SECURITY DEPOSIT",
            "",
            data.get("Security Deposit", ""),
            "",
            "6. MAINTENANCE",
            "",
            data.get("Maintenance", ""),
            "",
            "7. UTILITIES",
            "",
            data.get("Utilities", ""),
            "",
            "8. LATE PAYMENT",
            "",
            data.get("Late Fee", ""),
            "",
            "9. TERMINATION",
            "",
            data.get("Termination", ""),
            "",
            "10. ADDITIONAL TERMS",
            "",
            data.get("Additional Terms", "")
        ])

    elif document_type == "Service Agreement":

        lines.extend([
            "1. PARTIES",
            "",
            f"Service Provider: {data.get('Service Provider', '')}",
            f"Client: {data.get('Client', '')}",
            "",
            "2. SERVICE",
            "",
            f"Service / Project: {data.get('Service', '')}",
            f"Location: {data.get('Location', '')}",
            "",
            "3. SERVICE PERIOD",
            "",
            f"Start Date: {data.get('Start Date', '')}",
            f"Duration: {data.get('Duration', '')}",
            "",
            "4. SCOPE OF SERVICES",
            "",
            data.get("Scope", ""),
            "",
            "5. RESPONSIBILITIES",
            "",
            data.get("Responsibilities", ""),
            "",
            "6. PAYMENT",
            "",
            f"Service Fee: {data.get('Fee', '')}",
            f"Payment Schedule: {data.get('Payment Schedule', '')}",
            "",
            "7. INTELLECTUAL PROPERTY",
            "",
            data.get("Intellectual Property", ""),
            "",
            "8. TERMINATION",
            "",
            data.get("Termination", ""),
            "",
            "9. ADDITIONAL TERMS",
            "",
            data.get("Additional Terms", "")
        ])

    elif document_type == "Partnership Agreement":

        lines.extend([
            "1. PARTNERSHIP",
            "",
            f"Business Name: {data.get('Partnership Name', '')}",
            f"Partner 1: {data.get('Partner 1', '')}",
            f"Partner 2: {data.get('Partner 2', '')}",
            "",
            "2. NATURE OF BUSINESS",
            "",
            data.get("Business Activity", ""),
            "",
            "3. BUSINESS LOCATION",
            "",
            data.get("Business Location", ""),
            "",
            "4. CAPITAL CONTRIBUTION",
            "",
            f"Partner 1: {data.get('Capital Partner 1', '')}",
            f"Partner 2: {data.get('Capital Partner 2', '')}",
            "",
            "5. PROFIT AND LOSS SHARING",
            "",
            data.get("Profit Ratio", ""),
            "",
            "6. PARTNER RESPONSIBILITIES",
            "",
            data.get("Responsibilities", ""),
            "",
            "7. DECISION MAKING",
            "",
            data.get("Decision Making", ""),
            "",
            "8. WITHDRAWAL / RETIREMENT",
            "",
            data.get("Withdrawal", ""),
            "",
            "9. DISSOLUTION",
            "",
            data.get("Dissolution", ""),
            "",
            "10. ADDITIONAL TERMS",
            "",
            data.get("Additional Terms", "")
        ])

    lines.extend([
        "",
        "=" * 70,
        "SIGNATURES",
        "",
        "Party / Authorized Signatory: ______________________________",
        "",
        "Party / Authorized Signatory: ______________________________",
        "",
        "Date: ____________________"
    ])

    return "\n".join(lines)


# ============================================================
# GENERATE BUTTON
# ============================================================

st.markdown(
    "<div class='panel'>",
    unsafe_allow_html=True
)

st.subheader("Generate Document")

if st.button(
    "⚖️ Generate Legal Document",
    use_container_width=True
):

    generated_text = create_document(
        document_type,
        data
    )

    st.session_state["generated_document"] = generated_text
    st.session_state["generated_type"] = document_type

    st.success(
        f"{document_type} generated successfully."
    )

st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# DISPLAY GENERATED DOCUMENT
# ============================================================

if "generated_document" in st.session_state:

    st.markdown(
        "<div class='panel'>",
        unsafe_allow_html=True
    )

    st.subheader(
        f"Generated: {st.session_state['generated_type']}"
    )

    # Editable document
    edited_document = st.text_area(
        "Review / Edit Document",
        value=st.session_state["generated_document"],
        height=650
    )

    st.session_state["generated_document"] = edited_document

    # TXT download
    st.download_button(
        label="Download TXT",
        data=edited_document,
        file_name="LegalEase_Document.txt",
        mime="text/plain",
        use_container_width=True
    )

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    "<p style='text-align:center; color:#3b2417; "
    "font-weight:600;'>"
    "⚖️ LegalEase • Professional Legal Documentation"
    "</p>",
    unsafe_allow_html=True
)