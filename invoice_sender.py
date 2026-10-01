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
    page_title="Sheet Tools",
    page_icon="🛠️",
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
        <h1 style="color:#1a73e8; font-size: 48px;">🛠️ Sheet Tools</h1>
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
# CELL MAPPING (Simple Mode)
# ------------------------------------------------------------
CELL_MAP = {"q1": "Q1", "q3": "Q3"}

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
def get_last_sheet_prefill():
    try:
        sheets = get_spreadsheet().worksheets()
        if not sheets:
            return {}
        last = sheets[-1]
        prefill = {"sheet_name": last.title}
        for key, cell in CELL_MAP.items():
            try:
                prefill[key] = last.acell(cell).value or ""
            except Exception:
                prefill[key] = ""
        return prefill
    except Exception as e:
        st.warning(f"Could not read last sheet: {e}")
        return {}

@st.cache_data(ttl=15, show_spinner=False)
def read_range(sheet_name, cell_range):
    try:
        ws = get_spreadsheet().worksheet(sheet_name)
        data = ws.get(cell_range)
        if not data:
            return pd.DataFrame()
        max_cols = max(len(r) for r in data) if data else 0
        data = [(r + [""] * max_cols)[:max_cols] for r in data]
        df = pd.DataFrame(data)
        df = df.replace("", pd.NA).fillna("")
        return df
    except Exception as e:
        st.warning(f"Could not read {cell_range} from {sheet_name}: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=15, show_spinner=False)
def get_sheet_gid(sheet_name):
    try:
        ws = get_spreadsheet().worksheet(sheet_name)
        return ws.id
    except Exception:
        return None

# ------------------------------------------------------------
# WRITE HELPERS
# ------------------------------------------------------------
def duplicate_and_fill(new_name, data):
    ss = get_spreadsheet()
    sheets = ss.worksheets()
    if not sheets:
        raise Exception("No sheets found")

    source = sheets[-1]
    new_sheet = ss.duplicate_sheet(
        source_sheet_id=source.id,
        new_sheet_name=new_name,
        insert_sheet_index=len(sheets)
    )

    updates = []
    for key, cell in CELL_MAP.items():
        val = data.get(key, "")
        if val:
            updates.append({"range": cell, "values": [[val]]})

    if updates:
        new_sheet.batch_update(updates, value_input_option="USER_ENTERED")

    return new_sheet

def delete_sheet(sheet_name):
    ss = get_spreadsheet()
    ws = ss.worksheet(sheet_name)
    ss.del_worksheet(ws)

# ============================================================
# PDF GENERATION (ReportLab — plain)
# ============================================================
def build_pdf(sheet_name, df):
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
    title_style = ParagraphStyle("T", parent=styles["Title"], fontSize=14,
                                 leading=18, spaceAfter=6)
    small_style = ParagraphStyle("S", parent=styles["Normal"], fontSize=7, leading=9)

    elements = []
    elements.append(Paragraph(f"Invoice — {sheet_name}", title_style))
    elements.append(Paragraph(
        f"Generated: {_ist_now().strftime('%d-%b-%Y %H:%M')} IST", small_style))
    elements.append(Spacer(1, 6))

    table_data = []
    for _, row in df.iterrows():
        formatted = []
        for cell in row:
            s = str(cell) if cell is not None else ""
            if len(s) > 40:
                s = s[:37] + "..."
            formatted.append(Paragraph(s, small_style))
        table_data.append(formatted)

    if not table_data:
        elements.append(Paragraph("No data", styles["Normal"]))
    else:
        n = len(table_data[0])
        page_w = A4[0] - 20 * mm
        col_w = page_w / n
        table = Table(table_data, colWidths=[col_w] * n)
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
def send_email_with_pdf(to_email, subject, body, pdf_bytes, pdf_filename,
                         from_email, app_password):
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
# UI HEADER
# ============================================================
col_head1, col_head2 = st.columns([4, 1])
with col_head1:
    st.markdown(f"### 🛠️ Sheet Tools — Welcome, **{current_user.title()}**")
    st.caption(f"🕐 {_ist_now().strftime('%H:%M:%S')} IST")
with col_head2:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["authenticated"] = False
        st.rerun()

st.divider()

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
with st.sidebar:
    st.header("📊 Sheet Info")
    sheet_names = list_sheet_names()
    if sheet_names:
        st.caption(f"**{len(sheet_names)} sheets**")
        st.caption(f"Last sheet: **{sheet_names[-1]}**")
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

if not sheet_names:
    st.error("No sheets found. Verify spreadsheet ID and sharing.")
    st.stop()

# ============================================================
# TABS
# ============================================================
tab_dup, tab_pdf = st.tabs(["📋 Duplicate Sheet", "📄 Print PDF Invoice"])

# ============================================================
# TAB 1: DUPLICATE SHEET
# ============================================================
with tab_dup:
    st.subheader("📋 Create New Sheet from Last Sheet")
    st.info(f"📄 **Source sheet:** `{sheet_names[-1]}` — new sheet will be an exact copy with formatting preserved.")

    if "saved_prefill" not in st.session_state:
        st.session_state.saved_prefill = get_last_sheet_prefill()

    p = st.session_state.saved_prefill

    with st.form("dup_form", clear_on_submit=False):
        new_sheet_name = st.text_input(
            "Sheet Name *",
            value="",
            placeholder="Enter a name for the new sheet",
            help="Must be unique — cannot match an existing sheet"
        )

        st.markdown("---")
        col1, col2 = st.columns(2)
        with col1:
            q1_val = st.text_input("Value for Cell Q1", value=p.get("q1", ""),
                                   placeholder="Enter Q1 value")
        with col2:
            q3_val = st.text_input("Value for Cell Q3", value=p.get("q3", ""),
                                   placeholder="Enter Q3 value")

        submitted = st.form_submit_button("🆕 Create Sheet",
                                           use_container_width=True, type="primary")

    if submitted:
        if not new_sheet_name.strip():
            st.error("❌ Sheet name cannot be empty")
        elif new_sheet_name.strip() in sheet_names:
            st.error(f"❌ A sheet named '{new_sheet_name}' already exists")
        else:
            try:
                with st.spinner(f"Creating '{new_sheet_name}'..."):
                    duplicate_and_fill(new_sheet_name.strip(),
                                        {"q1": q1_val, "q3": q3_val})
                st.session_state.saved_prefill = {
                    "sheet_name": "", "q1": q1_val, "q3": q3_val,
                }
                st.cache_data.clear()
                st.success(f"✅ Sheet **'{new_sheet_name}'** created successfully!")
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"❌ Failed: {e}")

    # Existing sheets list
    st.divider()
    st.subheader("📚 Existing Sheets")
    sheet_df = pd.DataFrame({
        "#": range(1, len(sheet_names) + 1),
        "Sheet Name": sheet_names
    })
    st.dataframe(sheet_df, hide_index=True, use_container_width=True)

    with st.expander("🗑️ Delete a Sheet"):
        sheet_to_delete = st.selectbox("Select sheet",
                                        options=sheet_names,
                                        key="del_sheet")
        confirm = st.checkbox(f"Yes, delete **'{sheet_to_delete}'** permanently",
                              key="del_confirm")
        if st.button("Delete Sheet", type="secondary"):
            if confirm:
                try:
                    delete_sheet(sheet_to_delete)
                    st.cache_data.clear()
                    st.success(f"✅ Deleted '{sheet_to_delete}'")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Could not delete: {e}")
            else:
                st.warning("Please check the confirmation box")

# ============================================================
# TAB 2: PRINT PDF INVOICE
# ============================================================
with tab_pdf:
    st.subheader("📄 Invoice Preview & Print")

    col_pdf1, col_pdf2 = st.columns([3, 1])

    with col_pdf1:
        selected_sheet = st.selectbox(
            "Select Invoice Sheet",
            options=sheet_names,
            index=len(sheet_names) - 1,
            key="pdf_sheet_select"
        )

    with col_pdf2:
        cell_range = st.text_input("Cell Range", value="A1:N44", key="pdf_range")

    with st.spinner(f"Loading {cell_range}..."):
        invoice_df = read_range(selected_sheet, cell_range)

    if invoice_df.empty:
        st.warning("No data in the selected range")
    else:
        st.caption(f"Showing **{len(invoice_df)} rows × {len(invoice_df.columns)} columns**")
        st.dataframe(invoice_df, use_container_width=True, height=400, hide_index=True)

        # =====================================================
        # BUILD URLS
        # =====================================================
        try:
            ss = get_spreadsheet()
            sheet_ws = ss.worksheet(selected_sheet)
            sheet_gid = sheet_ws.id
            spreadsheet_id = st.secrets["spreadsheet_id"]

            encoded_range = cell_range.replace(":", "%3A")

            pdf_export_url = (
                f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
                f"?format=pdf"
                f"&size=A4"
                f"&portrait=true"
                f"&fitw=true"
                f"&gridlines=false"
                f"&printtitle=false"
                f"&sheetnames=false"
                f"&pagenum=UNDEFINED"
                f"&horizontal_alignment=CENTER"
                f"&vertical_alignment=TOP"
                f"&fzr=true"
                f"&fzc=true"
                f"&gid={sheet_gid}"
                f"&range={encoded_range}"
            )

            edit_url = (
                f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
                f"/edit#gid={sheet_gid}&range={cell_range}"
            )

        except Exception as e:
            st.error(f"Could not build URLs: {e}")
            st.stop()

        # =====================================================
        # ANDROID PRINT WORKFLOW
        # =====================================================
        st.divider()
        st.markdown("### 📱 Print from Android — 3 Easy Steps")

        st.markdown("""
        <div style="background: #e8f5e9; padding: 12px 16px; border-radius: 10px;
                    border-left: 5px solid #4caf50; margin-bottom: 12px;">
            <b style="color: #2e7d32;">Step 1 — Open the PDF</b><br>
            <span style="color: #555; font-size: 13px;">
                Tap the button below. The PDF opens in a new tab with all
                formatting, colors, and images intact.
            </span>
        </div>
        """, unsafe_allow_html=True)

        st.link_button(
            "🖨️ Open Pixel-Perfect PDF (matches Ctrl+P)",
            pdf_export_url,
            use_container_width=True,
            type="primary"
        )

        st.markdown("""
        <div style="background: #fff3e0; padding: 12px 16px; border-radius: 10px;
                    border-left: 5px solid #ff9800; margin-top: 16px; margin-bottom: 12px;">
            <b style="color: #e65100;">Step 2 — In Chrome, tap Share → Print</b><br>
            <span style="color: #555; font-size: 13px;">
                The PDF opens in Chrome. Look for the <b>Share icon</b> at the top
                or tap the <b>⋮ menu → Share → Print</b>.
            </span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="background: #e3f2fd; padding: 12px 16px; border-radius: 10px;
                    border-left: 5px solid #2196f3;">
            <b style="color: #0d47a1;">Step 3 — Choose your printer</b><br>
            <span style="color: #555; font-size: 13px;">
                Android's print dialog appears. Select your printer and print.
            </span>
        </div>
        """, unsafe_allow_html=True)

        # =====================================================
        # OTHER OPTIONS
        # =====================================================
        st.divider()
        st.markdown("### 🔄 Other Options")

        col_alt1, col_alt2 = st.columns(2)

        with col_alt1:
            st.link_button(
                "📊 Open in Google Sheets",
                edit_url,
                use_container_width=True
            )
            st.caption("Edit the sheet before printing")

        with col_alt2:
            st.link_button(
                "📥 Download PDF Directly",
                pdf_export_url,
                use_container_width=True
            )
            st.caption("Saves to Downloads folder")

        # =====================================================
        # INSTRUCTIONS
        # =====================================================
        with st.expander("📖 Detailed Android Print Instructions"):
            st.markdown("""
            ### How to Print the Invoice from Android

            **Option A — Print Directly from Chrome**
            1. Tap **"🖨️ Open Pixel-Perfect PDF"** above
            2. Chrome opens the PDF in a new tab
            3. Tap the **Share icon** (📤) at the top right  
               *(or tap ⋮ menu → Share)*
            4. Tap **Print**
            5. Android's print dialog appears
            6. Select your printer (or "Save as PDF")
            7. Tap **Print**

            **Option B — Save Then Print**
            1. Tap **"📥 Download PDF Directly"** above
            2. The PDF downloads to your **Downloads** folder
            3. Open **Files** app → Downloads → tap the PDF
            4. Tap the **⋮ menu → Print**
            5. Select printer → Print

            **Option C — Send via Email Then Print**
            1. Open the PDF using Option A
            2. Tap **Share** → **Gmail**
            3. Send to your own email
            4. Open the email on your phone
            5. Tap the attachment → Print

            ---

            ### Why 3 Taps?
            Android and Chrome don't allow any web app to open the print dialog
            automatically — for security. The 3-tap workflow above is the
            minimum possible on Android.
            """)

        # =====================================================
        # REPORTLAB PDF + EMAIL
        # =====================================================
        st.divider()
        st.markdown("### 📧 Send PDF by Email")

        st.caption(
            "Sends a **plain-text PDF** (no images/colors) as an email attachment. "
            "Use this when you need to programmatically email the invoice."
        )

        try:
            default_email = st.secrets.get("DEFAULT_EMAIL", "")
            smtp_email = st.secrets.get("SMTP_EMAIL", "")
            smtp_password = st.secrets.get("SMTP_APP_PASSWORD", "")
        except Exception:
            default_email = ""
            smtp_email = ""
            smtp_password = ""

        with st.form("pdf_form"):
            col_a, col_b = st.columns([2, 1])

            with col_a:
                to_email = st.text_input("Recipient Email", value=default_email,
                                         placeholder="name@example.com")
                subject = st.text_input(
                    "Subject",
                    value=f"Invoice — {selected_sheet} — {_ist_now().strftime('%d-%b-%Y')}"
                )

            with col_b:
                pdf_filename = st.text_input("PDF Filename",
                                              value=f"{selected_sheet}.pdf")

            message = st.text_area(
                "Message",
                value=f"Please find attached the invoice.\n\nGenerated on "
                      f"{_ist_now().strftime('%d-%b-%Y %H:%M')} IST",
                height=80
            )

            col_send, col_download = st.columns([1, 1])

            with col_send:
                send_clicked = st.form_submit_button("📧 Send PDF to Email",
                                                      use_container_width=True,
                                                      type="primary")

            with col_download:
                gen_download = st.form_submit_button("📥 Generate PDF for Download",
                                                      use_container_width=True)

        if gen_download:
            with st.spinner("Generating PDF..."):
                try:
                    pdf_bytes = build_pdf(selected_sheet, invoice_df)
                    st.session_state["pdf_bytes"] = pdf_bytes
                    st.session_state["pdf_filename"] = pdf_filename
                    st.success("✅ PDF ready")
                except Exception as e:
                    st.error(f"❌ Failed: {e}")

        if "pdf_bytes" in st.session_state:
            st.download_button(
                "📥 Click to Download ReportLab PDF",
                data=st.session_state["pdf_bytes"],
                file_name=st.session_state.get("pdf_filename", "invoice.pdf"),
                mime="application/pdf",
                use_container_width=True
            )

        if send_clicked:
            if not is_valid_email(to_email):
                st.error("❌ Invalid email address")
            elif not smtp_email or not smtp_password:
                st.error("❌ SMTP credentials missing in secrets")
            else:
                with st.spinner(f"Sending to {to_email}..."):
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
                        st.success(f"✅ Sent to **{to_email}**")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Failed: {e}")

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.divider()
st.caption("🛠️ Sheet Tools · Built by S. Mohapatra")
