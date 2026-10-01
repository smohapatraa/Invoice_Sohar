import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
import io
import re

# PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import mm

# Email sending
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Invoice Sender",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ------------------------------------------------------------
# IST HELPERS
# ------------------------------------------------------------
def _ist_now():
    ist = timezone(timedelta(hours=5, minutes=30))
    return datetime.now(ist)

# ============================================================
# LOGIN
# ============================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div style="text-align:center; padding: 40px 0;">
        <h1 style="color:#1a73e8; font-size: 48px;">📧 Invoice Sender</h1>
        <p style="color:#888; font-size: 16px;">Please log in to continue</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            u = st.text_input("Username", placeholder="Enter your username")
            p = st.text_input("Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("🔐 Log In", use_container_width=True, type="primary")

        if submitted:
            try:
                correct_user = st.secrets.get("MY_USERNAME", "")
                correct_pass = st.secrets.get("MY_PASSWORD", "")
            except Exception:
                correct_user = ""
                correct_pass = ""

            if u and p and u == correct_user and p == correct_pass:
                st.session_state.authenticated = True
                st.session_state.logged_in_user = u
                st.rerun()
            else:
                st.error("❌ Invalid username or password")

    st.stop()

current_user = st.session_state.get("logged_in_user", "user")

# ============================================================
# GOOGLE SHEETS
# ============================================================
@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_info = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    return get_gspread_client().open_by_key(st.secrets["spreadsheet_id"])

# ------------------------------------------------------------
# READ HELPERS
# ------------------------------------------------------------
@st.cache_data(ttl=15, show_spinner=False)
def list_sheet_names():
    try:
        return [ws.title for ws in get_spreadsheet().worksheets()]
    except Exception as e:
        st.error(f"Could not fetch sheets: {e}")
        return []

@st.cache_data(ttl=15, show_spinner=False)
def read_range(sheet_name, cell_range="A1:N44"):
    try:
        ws = get_spreadsheet().worksheet(sheet_name)
        data = ws.get(cell_range)
        if not data:
            return pd.DataFrame()
        # Normalize row width
        max_cols = max(len(r) for r in data) if data else 0
        data = [(r + [""] * max_cols)[:max_cols] for r in data]
        df = pd.DataFrame(data)
        df = df.replace("", pd.NA).fillna("")
        return df
    except Exception as e:
        st.warning(f"Could not read {cell_range} from {sheet_name}: {e}")
        return pd.DataFrame()

# ============================================================
# PDF GENERATION
# ============================================================
def build_pdf(sheet_name, df, start_range="A1"):
    """Build a PDF from a DataFrame of cell values."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        leftMargin=10 * mm,
        rightMargin=10 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Title"],
        fontSize=14,
        leading=18,
        spaceAfter=6,
    )
    small_style = ParagraphStyle(
        "SmallStyle",
        parent=styles["Normal"],
        fontSize=7,
        leading=9,
    )

    elements = []
    elements.append(Paragraph(f"Invoice — {sheet_name}", title_style))
    elements.append(Paragraph(f"Generated: {_ist_now().strftime('%d-%b-%Y %H:%M')} IST", small_style))
    elements.append(Spacer(1, 6))

    # Prepare table data
    table_data = []
    for _, row in df.iterrows():
        formatted_row = []
        for cell in row:
            cell_str = str(cell) if cell is not None else ""
            # Truncate long values
            if len(cell_str) > 40:
                cell_str = cell_str[:37] + "..."
            formatted_row.append(Paragraph(cell_str, small_style))
        table_data.append(formatted_row)

    if not table_data:
        elements.append(Paragraph("No data available.", styles["Normal"]))
    else:
        # Auto column widths
        n_cols = len(table_data[0])
        page_width = A4[0] - 20 * mm
        col_width = page_width / n_cols

        table = Table(table_data, colWidths=[col_width] * n_cols)
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ]))
        elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
# EMAIL SENDING
# ============================================================
def send_email_with_pdf(to_email, subject, body, pdf_bytes, pdf_filename, from_email, app_password):
    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    part = MIMEApplication(pdf_bytes, _subtype="pdf")
    part.add_header("Content-Disposition", f'attachment; filename="{pdf_filename}"')
    msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(from_email, app_password)
        server.send_message(msg)

def is_valid_email(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) is not None

# ============================================================
# UI
# ============================================================
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown(f"### 📧 Invoice Sender — Welcome, **{current_user.title()}**")
    st.caption(f"🕐 {_ist_now().strftime('%H:%M:%S')} IST")
with col_head2:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

st.divider()

# ------------------------------------------------------------
# SIDEBAR — Sheet + Range
# ------------------------------------------------------------
with st.sidebar:
    st.header("📊 Sheet Config")

    sheet_names = list_sheet_names()
    if sheet_names:
        default_index = len(sheet_names) - 1
        selected_sheet = st.selectbox(
            "Select Invoice Sheet",
            options=sheet_names,
            index=default_index,
            key="sheet_select"
        )
    else:
        st.error("No sheets found")
        st.stop()

    cell_range = st.text_input("Cell Range", value="A1:N44", key="range_input")

    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# ------------------------------------------------------------
# LOAD DATA
# ------------------------------------------------------------
with st.spinner(f"Loading {cell_range} from '{selected_sheet}'..."):
    invoice_df = read_range(selected_sheet, cell_range)

# ------------------------------------------------------------
# PREVIEW
# ------------------------------------------------------------
st.subheader(f"📄 Invoice Preview — {selected_sheet}")

if invoice_df.empty:
    st.warning("No data in this range")
    st.stop()

# Style the dataframe
styled_df = invoice_df.style.set_properties(**{
    "font-size": "12px",
    "padding": "4px",
    "border-color": "#ddd",
})

st.dataframe(
    invoice_df,
    use_container_width=True,
    height=600,
    hide_index=True
)

st.caption(f"Showing **{len(invoice_df)} rows × {len(invoice_df.columns)} columns** from `{cell_range}`")

# ------------------------------------------------------------
# SEND OPTIONS
# ------------------------------------------------------------
st.divider()
st.subheader("📧 Send as PDF")

# Default email + credentials
try:
    default_email = st.secrets.get("DEFAULT_EMAIL", "")
    smtp_email = st.secrets.get("SMTP_EMAIL", "")
    smtp_password = st.secrets.get("SMTP_APP_PASSWORD", "")
except Exception:
    default_email = ""
    smtp_email = ""
    smtp_password = ""

with st.form("send_form"):
    col_a, col_b = st.columns([2, 1])

    with col_a:
        to_email = st.text_input(
            "Recipient Email",
            value=default_email,
            placeholder="name@example.com"
        )
        subject = st.text_input(
            "Subject",
            value=f"Invoice — {selected_sheet} — {_ist_now().strftime('%d-%b-%Y')}"
        )

    with col_b:
        pdf_filename = st.text_input(
            "PDF Filename",
            value=f"{selected_sheet}.pdf"
        )

    message = st.text_area(
        "Message",
        value=f"Please find attached the invoice from {selected_sheet}.\n\nGenerated on {_ist_now().strftime('%d-%b-%Y %H:%M')} IST",
        height=100
    )

    col_send, col_download = st.columns([1, 1])

    with col_send:
        send_clicked = st.form_submit_button(
            "📧 Send PDF to Email",
            use_container_width=True,
            type="primary"
        )

    with col_download:
        generate_for_download = st.form_submit_button(
            "📥 Generate PDF (Download)",
            use_container_width=True
        )

# ------------------------------------------------------------
# HANDLE: GENERATE FOR DOWNLOAD
# ------------------------------------------------------------
if generate_for_download:
    with st.spinner("Generating PDF..."):
        try:
            pdf_bytes = build_pdf(selected_sheet, invoice_df)
            st.session_state["pdf_bytes"] = pdf_bytes
            st.session_state["pdf_filename"] = pdf_filename
            st.success("✅ PDF generated")
        except Exception as e:
            st.error(f"❌ PDF generation failed: {e}")

if "pdf_bytes" in st.session_state:
    st.download_button(
        "📥 Click to Download PDF",
        data=st.session_state["pdf_bytes"],
        file_name=st.session_state.get("pdf_filename", "invoice.pdf"),
        mime="application/pdf",
        use_container_width=True
    )

# ------------------------------------------------------------
# HANDLE: SEND EMAIL
# ------------------------------------------------------------
if send_clicked:
    if not is_valid_email(to_email):
        st.error("❌ Please enter a valid email address")
    elif not smtp_email or not smtp_password:
        st.error("❌ SMTP credentials not configured in secrets. Add SMTP_EMAIL and SMTP_APP_PASSWORD.")
    else:
        with st.spinner(f"Sending PDF to {to_email}..."):
            try:
                pdf_bytes = build_pdf(selected_sheet, invoice_df)
                send_email_with_pdf(
                    to_email=to_email,
                    subject=subject,
                    body=message,
                    pdf_bytes=pdf_bytes,
                    pdf_filename=pdf_filename,
                    from_email=smtp_email,
                    app_password=smtp_password,
                )
                st.success(f"✅ PDF sent to **{to_email}**")
                st.balloons()
            except Exception as e:
                st.error(f"❌ Failed to send email: {e}")

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.divider()
st.caption("📧 Invoice Sender · Built by S. Mohapatra")
