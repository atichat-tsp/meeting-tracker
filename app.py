import streamlit as st
import json
import requests
import pandas as pd
import anthropic
import os
import io
import time
import base64
from datetime import datetime
import pytz

# =============================================
# 1. ตั้งค่าหน้าตาเว็บ
# =============================================
st.set_page_config(page_title="TSP Meeting Tracker", layout="wide")

# =============================================
# 2. โหลด Config
# =============================================
try:
    line_bot_token   = st.secrets["LINE_BOT_TOKEN"]
    line_user_id     = st.secrets["LINE_USER_ID"]
    line_group_id    = st.secrets.get("LINE_GROUP_ID", "")
    anthropic_key    = st.secrets["ANTHROPIC_API_KEY"]
    tenant_id        = st.secrets["AZURE_TENANT_ID"]
    client_id        = st.secrets["AZURE_CLIENT_ID"]
    client_secret    = st.secrets["AZURE_CLIENT_SECRET"]
    onedrive_folder  = st.secrets.get("ONEDRIVE_FOLDER", "Meeting Tracker")
    user_email       = st.secrets.get("ONEDRIVE_USER_EMAIL", "atichat@tspmetal.com")
    email_recipients = st.secrets.get("EMAIL_RECIPIENTS", "teetat@tspmetal.com,samorn@tspmetal.com,dr.witjun@gmail.com,warun@tspmetal.com,panuwat@tspmetal.com")
except Exception:
    line_bot_token   = os.environ.get("LINE_BOT_TOKEN", "")
    line_user_id     = os.environ.get("LINE_USER_ID", "")
    line_group_id    = os.environ.get("LINE_GROUP_ID", "")
    anthropic_key    = os.environ.get("ANTHROPIC_API_KEY", "")
    tenant_id        = os.environ.get("AZURE_TENANT_ID", "")
    client_id        = os.environ.get("AZURE_CLIENT_ID", "")
    client_secret    = os.environ.get("AZURE_CLIENT_SECRET", "")
    onedrive_folder  = os.environ.get("ONEDRIVE_FOLDER", "Meeting Tracker")
    user_email       = os.environ.get("ONEDRIVE_USER_EMAIL", "atichat@tspmetal.com")
    email_recipients = os.environ.get("EMAIL_RECIPIENTS", "teetat@tspmetal.com")

# =============================================
# 3. Header
# =============================================
TH_TZ = pytz.timezone("Asia/Bangkok")
now = datetime.now(TH_TZ)
THAI_DAYS = ["จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์","อาทิตย์"]
THAI_MONTHS = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
day_name  = THAI_DAYS[now.weekday()]
day_num   = now.day
month_name = THAI_MONTHS[now.month]
year_thai = now.year + 543
time_str  = now.strftime("%H:%M")

logo_b64 = ""
logo_path = os.path.join(os.path.dirname(__file__), "logo.png")
if os.path.exists(logo_path):
    with open(logo_path, "rb") as f:
        logo_b64 = base64.b64encode(f.read()).decode()

logo_img_html = f"<img src='data:image/png;base64,{logo_b64}' style='width:56px;height:56px;object-fit:contain;border-radius:8px;'/>" if logo_b64 else "<div style='width:56px;height:56px;background:#1a5c3a;border-radius:8px;'></div>"

st.markdown(f"""
<div style="padding:8px 0 16px;">
  <div style="background:#f0f7f4; border:0.5px solid #b2d8c8; border-radius:12px; padding:14px 24px 0;">
    <div style="display:grid; grid-template-columns:64px 1fr auto; align-items:center; gap:16px;">
      <div>{logo_img_html}</div>
      <div style="text-align:center;">
        <div style="font-size:28px; font-weight:500; color:#1a5c3a; line-height:1.2;">TSP Metal Works</div>
        <div style="font-size:17px; color:#2d7a56; margin-top:4px; font-weight:500;">บันทึกติดตามและสรุปรายงานการประชุม</div>
      </div>
      <div style="text-align:right;">
        <div style="font-size:13px; font-weight:500; color:#1a5c3a;">วัน{day_name}ที่ {day_num} {month_name} {year_thai}</div>
        <div style="font-size:12px; color:#2d7a56;">เวลา {time_str} น.</div>
      </div>
    </div>
    <div style="margin-top:12px; height:3px; background:linear-gradient(90deg, #1a5c3a 0%, #2d9e6b 50%, #1a5c3a 100%); border-radius:2px 2px 0 0;"></div>
  </div>
</div>
""", unsafe_allow_html=True)

# =============================================
# 4. ฟังก์ชัน OneDrive
# =============================================
def get_access_token():
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    resp = requests.post(url, data=data, timeout=15)
    resp.raise_for_status()
    return resp.json()["access_token"]

def upload_to_onedrive(file_bytes: bytes, filename: str) -> str:
    try:
        token = get_access_token()
        upload_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/drive/root:/{onedrive_folder}/{filename}:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        resp = requests.put(upload_url, headers=headers, data=file_bytes, timeout=30)
        resp.raise_for_status()
        return resp.json().get("webUrl", "")
    except Exception as e:
        st.error(f"❌ Upload OneDrive ไม่สำเร็จ: {e}")
        return ""

def download_from_onedrive(filename: str) -> bytes:
    try:
        token = get_access_token()
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/drive/root:/{onedrive_folder}/{filename}:/content"
        headers = {"Authorization": f"Bearer {token}"}
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 200:
            return resp.content
        return b""
    except Exception:
        return b""

# =============================================
# 5. ฟังก์ชัน Email
# =============================================
def send_email_with_attachment(file_bytes: bytes, subject: str, body: str):
    """ส่ง Email พร้อมไฟล์แนบ — ตรวจสอบประเภทไฟล์อัตโนมัติ (PDF หรือ HTML fallback)"""
    try:
        token = get_access_token()
        recipients = [r.strip() for r in email_recipients.split(",")]
        to_list = [{"emailAddress": {"address": r}} for r in recipients if r]
        file_b64 = base64.b64encode(file_bytes).decode()

        # ตรวจสอบประเภทไฟล์จาก magic bytes
        is_real_pdf = file_bytes[:4] == b'%PDF'
        if is_real_pdf:
            filename    = f"meeting_report_{now.strftime('%Y%m%d')}.pdf"
            content_type = "application/pdf"
        else:
            # WeasyPrint ไม่ติดตั้ง — ส่งเป็น HTML แทน (เปิดได้ทุก browser)
            filename    = f"meeting_report_{now.strftime('%Y%m%d')}.html"
            content_type = "text/html"

        message = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body},
                "toRecipients": to_list,
                "attachments": [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentType": content_type,
                    "contentBytes": file_b64
                }]
            }
        }
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        resp = requests.post(url, headers=headers, json=message, timeout=30)
        return resp.status_code == 202, is_real_pdf
    except Exception as e:
        st.error(f"❌ ส่ง Email ไม่สำเร็จ: {e}")
        return False, False

# Alias เพื่อ backward-compat
def send_email_with_pdf(pdf_bytes: bytes, subject: str, body: str):
    ok, _ = send_email_with_attachment(pdf_bytes, subject, body)
    return ok

# =============================================
# 6. ฟังก์ชัน LINE
# =============================================
def send_line_message(message: str, target_id: str):
    if not line_bot_token or not target_id:
        return
    url = "https://api.line.me/v2/bot/message/push"
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {line_bot_token}"}
    payload = {"to": target_id, "messages": [{"type": "text", "text": message}]}
    try:
        requests.post(url, headers=headers, json=payload, timeout=10)
    except Exception:
        pass

# =============================================
# 7. ฟังก์ชันสร้าง Excel
# =============================================
def build_excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            max_len = max((len(str(cell.value or '')) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)
    return buf.getvalue()

# =============================================
# 8. ลงทะเบียน THSarabun / Sarabun font สำหรับ reportlab
# =============================================
import os as _os

def _register_thai_font():
    """ลงทะเบียน THSarabun font จาก local directory หรือ system fonts"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    # หา font directory ที่อยู่ข้างๆ app.py
    base_dir   = _os.path.dirname(_os.path.abspath(__file__))
    fonts_dirs = [
        _os.path.join(base_dir, "fonts"),          # ./fonts/ (ติด deploy ไปด้วย)
        _os.path.join(base_dir),                   # root ของโปรเจกต์
        "/usr/share/fonts/truetype/tlwg",          # Laksaman ที่ติดตั้งแล้ว
    ]

    # ชื่อไฟล์ที่ยอมรับ (ตามลำดับความสำคัญ)
    candidates_normal = [
        "THSarabunNew.ttf", "THSarabun.ttf",
        "Sarabun-Regular.ttf", "Laksaman.ttf",
    ]
    candidates_bold = [
        "THSarabunNew-Bold.ttf", "THSarabun-Bold.ttf",
        "Sarabun-Bold.ttf", "Laksaman-Bold.ttf",
    ]

    def find_font(candidates):
        for d in fonts_dirs:
            for name in candidates:
                p = _os.path.join(d, name)
                if _os.path.exists(p):
                    return p
        return None

    normal_path = find_font(candidates_normal)
    bold_path   = find_font(candidates_bold)

    if normal_path:
        pdfmetrics.registerFont(TTFont("THSarabun",      normal_path))
        pdfmetrics.registerFont(TTFont("THSarabun-Bold", bold_path or normal_path))
        pdfmetrics.registerFontFamily("THSarabun",
            normal="THSarabun", bold="THSarabun-Bold",
            italic="THSarabun", boldItalic="THSarabun-Bold")
        return True
    return False

_THAI_FONT_OK = _register_thai_font()
_THAI_FONT    = "THSarabun" if _THAI_FONT_OK else "Helvetica"
_THAI_FONT_B  = "THSarabun-Bold" if _THAI_FONT_OK else "Helvetica-Bold"

# =============================================
# 9. ฟังก์ชันสร้าง PDF ด้วย reportlab + THSarabun
# =============================================
def build_pdf_bytes(df_active: pd.DataFrame, df_history: pd.DataFrame, updates: dict) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    )
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

    COLOR_GREEN      = colors.HexColor("#1a5c3a")
    COLOR_GREEN_LITE = colors.HexColor("#e8f5ee")
    COLOR_ROW_ODD    = colors.HexColor("#f4f9f6")
    COLOR_BORDER     = colors.HexColor("#b2d8c8")

    # ─── Styles ────────────────────────────────────────────────────────────────
    s_title   = ParagraphStyle("title",   fontName=_THAI_FONT_B, fontSize=18,
                                textColor=COLOR_GREEN, alignment=TA_CENTER, spaceAfter=2)
    s_sub     = ParagraphStyle("sub",     fontName=_THAI_FONT,   fontSize=13,
                                textColor=colors.HexColor("#2d7a56"), alignment=TA_CENTER, spaceAfter=2)
    s_date    = ParagraphStyle("date",    fontName=_THAI_FONT,   fontSize=11,
                                textColor=colors.HexColor("#555555"), alignment=TA_RIGHT, spaceAfter=8)
    s_section = ParagraphStyle("section", fontName=_THAI_FONT_B, fontSize=13,
                                textColor=COLOR_GREEN, spaceBefore=12, spaceAfter=4)
    s_none    = ParagraphStyle("none",    fontName=_THAI_FONT,   fontSize=11,
                                textColor=colors.HexColor("#888888"), leftIndent=6)
    s_cell    = ParagraphStyle("cell",    fontName=_THAI_FONT,   fontSize=9,
                                leading=12, wordWrap="CJK")
    s_hdr     = ParagraphStyle("hdr",     fontName=_THAI_FONT_B, fontSize=9,
                                textColor=colors.white, leading=12, wordWrap="CJK")

    def safe(v):
        return str(v) if v is not None and str(v) not in ("nan", "None", "") else "-"

    def make_table(df, cols_wanted):
        """สร้าง reportlab Table จาก DataFrame"""
        existing = [c for c in cols_wanted if c in df.columns]
        if not existing:
            return None

        # Header row
        header = [Paragraph(c, s_hdr) for c in existing]
        rows   = [header]

        for i, (_, row) in enumerate(df.iterrows()):
            r = [Paragraph(safe(row.get(c, "")), s_cell) for c in existing]
            rows.append(r)

        # คำนวณความกว้างคอลัมน์
        page_w   = A4[0] - 28*mm  # margin รวม
        col_w_map = {
            "Issue ID":    18*mm,
            "หัวข้อปัญหา": None,   # stretch
            "ผู้แจ้ง":      22*mm,
            "Priority":    16*mm,
            "ผู้รับผิดชอบ": 24*mm,
            "สถานะ":        20*mm,
            "กำหนดเสร็จ":  22*mm,
            "วันที่แจ้ง":   22*mm,
            "อัปเดตล่าสุด": 26*mm,
            "งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม": None,
            "ความคิดเห็น / Comments": None,
        }
        fixed  = sum(v for c, v in col_w_map.items() if c in existing and v is not None)
        n_flex = sum(1  for c    in existing if col_w_map.get(c) is None)
        flex_w = (page_w - fixed) / max(n_flex, 1)
        col_widths = [col_w_map.get(c) or flex_w for c in existing]

        t = Table(rows, colWidths=col_widths, repeatRows=1)
        style = TableStyle([
            # Header
            ("BACKGROUND",  (0, 0), (-1, 0), COLOR_GREEN),
            ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
            ("FONTNAME",    (0, 0), (-1, 0), _THAI_FONT_B),
            ("FONTSIZE",    (0, 0), (-1, 0), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COLOR_ROW_ODD, colors.white]),
            ("GRID",        (0, 0), (-1, -1), 0.4, COLOR_BORDER),
            ("VALIGN",      (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",  (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING",(0, 0), (-1, -1), 5),
        ])
        t.setStyle(style)
        return t

    # ─── Build story ──────────────────────────────────────────────────────────
    buf  = io.BytesIO()
    doc  = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=14*mm,  bottomMargin=14*mm,
    )
    story = []

    # หัวรายงาน
    story.append(Paragraph("TSP Metal Works", s_title))
    story.append(Paragraph("บันทึกติดตามและสรุปรายงานการประชุม", s_sub))
    story.append(Paragraph(f"วัน{day_name}ที่ {day_num} {month_name} {year_thai} เวลา {time_str} น.", s_date))
    story.append(HRFlowable(width="100%", thickness=2, color=COLOR_GREEN, spaceAfter=6))

    # สรุปตัวเลข
    summary_data = [
        [Paragraph("📌 งานคงค้าง (Active)", s_hdr),
         Paragraph(f"{len(df_active)} รายการ", s_hdr)],
        [Paragraph("✅ งานสิ้นสุดแล้ว (History)", s_hdr),
         Paragraph(f"{len(df_history)} รายการ", s_hdr)],
    ]
    summary_t = Table(summary_data, colWidths=[100*mm, 40*mm])
    summary_t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COLOR_GREEN_LITE),
        ("GRID",       (0, 0), (-1, -1), 0.4, COLOR_BORDER),
        ("FONTNAME",   (0, 0), (-1, -1), _THAI_FONT),
        ("FONTSIZE",   (0, 0), (-1, -1), 10),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
    ]))
    story.append(summary_t)
    story.append(Spacer(1, 8))

    # ตาราง Active
    cols_active = [
        "Issue ID", "หัวข้อปัญหา", "ผู้รับผิดชอบ", "Priority",
        "สถานะ", "กำหนดเสร็จ", "อัปเดตล่าสุด",
        "งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม",
    ]
    story.append(Paragraph("งานคงค้าง (Active Issues)", s_section))
    if df_active.empty:
        story.append(Paragraph("ไม่มีงานค้างครับ 🎉", s_none))
    else:
        t = make_table(df_active, cols_active)
        if t:
            story.append(t)
    story.append(Spacer(1, 6))

    # ตาราง History
    cols_hist = ["Issue ID", "หัวข้อปัญหา", "ผู้รับผิดชอบ", "สถานะ", "กำหนดเสร็จ", "วันที่แจ้ง"]
    story.append(Paragraph("งานที่สิ้นสุดแล้ว (History)", s_section))
    if df_history.empty:
        story.append(Paragraph("ยังไม่มีรายการ", s_none))
    else:
        t = make_table(df_history, cols_hist)
        if t:
            story.append(t)

    doc.build(story)
    return buf.getvalue()

EXPECTED_COLUMNS = [
    '#', 'Issue ID', 'หัวข้อปัญหา', 'ผู้แจ้ง', 'Priority',
    'ผู้รับผิดชอบ', 'สถานะ',
    'งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม',
    'ความคิดเห็น / Comments', 'กำหนดเสร็จ', 'วันที่แจ้ง', 'อัปเดตล่าสุด'
]

# =============================================
# 9. โหลดข้อมูลเก่าจาก OneDrive
# =============================================
@st.cache_data(ttl=60)
def load_active_issues():
    data = download_from_onedrive("Meeting_Issue_Log.xlsx")
    if data:
        try:
            df = pd.read_excel(io.BytesIO(data))
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=EXPECTED_COLUMNS)

# =============================================
# 10. UI — ส่วนที่ 1: อัปเดตงานค้าง
# =============================================
st.subheader("📋 ส่วนที่ 1 — อัปเดตงานคงค้าง")

df_existing = load_active_issues()
updates = {}

if df_existing.empty:
    st.info("ยังไม่มีงานค้างในระบบครับ")
else:
    status_options = ["In Progress", "Open", "Closed", "Done"]
    for _, row in df_existing.iterrows():
        if str(row.get('สถานะ', '')).lower() in ['closed', 'done']:
            continue
        issue_id    = row.get('Issue ID', '')
        topic       = row.get('หัวข้อปัญหา', '')
        owner       = row.get('ผู้รับผิดชอบ', '')
        status      = row.get('สถานะ', 'In Progress')
        due_date    = str(row.get('กำหนดเสร็จ', '-'))
        created_col = 'วันที่แจ้ง'
        created_date = str(row.get(created_col, now.strftime('%d/%m/%Y')))

        # คำนวณวันเกินกำหนด
        overdue_html = ""
        try:
            due_dt = datetime.strptime(due_date, '%d/%m/%Y').replace(tzinfo=TH_TZ)
            diff = (now - due_dt).days
            if diff > 0:
                overdue_html = f"<span style='color:#c0392b; font-weight:500;'>⚠️ เกินกำหนด {diff} วัน</span>"
            elif diff >= -3:
                overdue_html = f"<span style='color:#e67e22; font-weight:500;'>⏰ ใกล้ครบกำหนด {abs(diff)} วัน</span>"
        except Exception:
            pass

        with st.expander(f"**{issue_id}** — {topic} (ผู้รับผิดชอบ: {owner})", expanded=False):
            # แสดงวันที่
            st.markdown(f"""
<div style="display:flex; gap:24px; padding:6px 0 10px; font-size:13px; border-bottom:0.5px solid #eee; margin-bottom:10px;">
  <span>📅 <b>วันที่แจ้ง:</b> {created_date}</span>
  <span>🎯 <b>กำหนดเสร็จ:</b> {due_date}</span>
  {overdue_html}
</div>
""", unsafe_allow_html=True)

            col1, col2 = st.columns([3, 1])
            with col1:
                note = st.text_area(
                    "ความคืบหน้า / อุปสรรค / ติดที่ใคร:",
                    key=f"note_{issue_id}",
                    height=80,
                    placeholder="เช่น ติดต่อซัพพลายเออร์แล้ว รอใบเสนอราคา..."
                )
            with col2:
                new_status = st.selectbox(
                    "สถานะ:",
                    status_options,
                    index=status_options.index(status) if status in status_options else 0,
                    key=f"status_{issue_id}"
                )
            updates[issue_id] = {"note": note, "status": new_status}

# =============================================
# 11. UI — ส่วนที่ 2: บันทึกประชุมใหม่
# =============================================
st.divider()
st.subheader("✍️ ส่วนที่ 2 — บันทึกการประชุมใหม่")

raw_notes = st.text_area(
    "พิมพ์บันทึกประชุมดิบที่นี่:",
    height=200,
    placeholder="บอส: เรื่องกล้องวงจรปิดไปถึงไหนแล้ว?\nเกรียงศักดิ์: ติดต่อซัพพลายเออร์แล้วครับ รอใบเสนอราคา..."
)

# =============================================
# 12. ปุ่มประมวลผล
# =============================================
st.divider()
if st.button("🚀 ประมวลผล + อัปโหลด OneDrive + ส่ง Email + แจ้งเตือน LINE", type="primary"):

    if not any(v["note"].strip() for v in updates.values()) and not raw_notes.strip():
        st.warning("⚠️ กรุณากรอกข้อมูลอย่างน้อย 1 ส่วนครับ")
    else:
        with st.spinner("🤖 กำลังประมวลผล..."):
            try:
                # โหลดข้อมูลเก่า
                df_active  = load_active_issues()
                df_history = pd.DataFrame(columns=EXPECTED_COLUMNS)

                # อัปเดต Issue เก่า
                if not df_active.empty:
                    for issue_id, update_data in updates.items():
                        mask = df_active['Issue ID'] == issue_id
                        if not mask.any():
                            continue
                        if update_data["note"].strip():
                            old_note = str(df_active.loc[mask, 'งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม'].values[0])
                            new_note = f"{old_note}\n[{now.strftime('%d/%m/%Y %H:%M')}] {update_data['note']}".strip()
                            df_active.loc[mask, 'งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม'] = new_note
                        df_active.loc[mask, 'สถานะ'] = update_data["status"]
                        df_active.loc[mask, 'อัปเดตล่าสุด'] = now.strftime('%d/%m/%Y %H:%M')

                    # แยก Closed ออกไป History
                    done_mask  = df_active['สถานะ'].str.lower().isin(['closed', 'done'])
                    df_history = df_active[done_mask].copy()
                    df_active  = df_active[~done_mask].copy()

                # ประมวลผลบันทึกใหม่ด้วย Claude
                if raw_notes.strip():
                    # หาเลข Issue ล่าสุด
                    last_num = 0
                    all_issues = pd.concat([df_active, df_history], ignore_index=True)
                    if not all_issues.empty and 'Issue ID' in all_issues.columns:
                        nums = all_issues['Issue ID'].str.extract(r'(\d+)').dropna()[0].astype(int)
                        if len(nums) > 0:
                            last_num = nums.max()

                    prompt = f"""
คุณคือผู้ช่วยเลขานุการ วิเคราะห์บันทึกประชุมต่อไปนี้แล้วแยก Action Items ใหม่เป็น JSON Array
เริ่มเลข Issue ID จาก ISS-{last_num+1:03d} ไปเรื่อยๆ
(ตอบเป็น JSON เท่านั้น):

คอลัมน์: "Issue ID","หัวข้อปัญหา","ผู้แจ้ง","Priority","ผู้รับผิดชอบ","สถานะ","งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม","ความคิดเห็น / Comments","กำหนดเสร็จ"

กฎ:
- ผู้ติดตาม = "Teetat" เสมอ
- สถานะ: เสร็จ/ปิด → "Closed" | อื่นๆ → "In Progress"
- กำหนดเสร็จ: DD/MM/YYYY ปี 2026

บันทึกประชุม:
\"\"\"{raw_notes}\"\"\"
"""
                    client = anthropic.Anthropic(api_key=anthropic_key)
                    for attempt in range(3):
                        try:
                            response = client.messages.create(
                                model="claude-sonnet-4-6",
                                max_tokens=8192,
                                messages=[{"role": "user", "content": prompt}]
                            )
                            break
                        except Exception:
                            if attempt < 2:
                                time.sleep(3)
                            else:
                                raise

                    raw_text = response.content[0].text.strip()
                    if raw_text.startswith("```"):
                        raw_text = raw_text.split("```")[1]
                        if raw_text.startswith("json"):
                            raw_text = raw_text[4:]
                    raw_text = raw_text.strip()

                    new_items = json.loads(raw_text)
                    df_new = pd.DataFrame(new_items)
                    df_new.insert(0, '#', range(len(df_active)+1, len(df_active)+len(df_new)+1))
                    if len(df_new.columns) == len(EXPECTED_COLUMNS) - 2:
                        df_new.columns = EXPECTED_COLUMNS[:-2]
                    df_new['วันที่แจ้ง'] = now.strftime('%d/%m/%Y')
                    df_new['อัปเดตล่าสุด'] = now.strftime('%d/%m/%Y %H:%M')
                    for col in EXPECTED_COLUMNS:
                        if col not in df_new.columns:
                            df_new[col] = ''
                    df_new = df_new[EXPECTED_COLUMNS]

                    new_done = df_new['สถานะ'].str.lower().isin(['closed', 'done'])
                    df_history = pd.concat([df_history, df_new[new_done]], ignore_index=True)
                    df_active  = pd.concat([df_active,  df_new[~new_done]], ignore_index=True)
                    df_active['#']  = range(1, len(df_active)+1)
                    df_history['#'] = range(1, len(df_history)+1)

                # แสดงผล
                st.subheader(f"📌 งานคงค้าง (Active) — {len(df_active)} รายการ")
                if not df_active.empty:
                    st.dataframe(df_active, use_container_width=True)
                else:
                    st.info("ไม่มีงานค้างครับ 🎉")

                st.subheader(f"📜 งานที่สิ้นสุดแล้ว (History) — {len(df_history)} รายการ")
                if not df_history.empty:
                    st.dataframe(df_history, use_container_width=True)

                # สร้าง Excel
                active_bytes  = build_excel_bytes(df_active,  'Active Issues')
                history_bytes = build_excel_bytes(df_history, 'History Log')

                # Upload OneDrive
                active_link = history_link = ""
                with st.spinner("☁️ อัปโหลดขึ้น OneDrive..."):
                    active_link  = upload_to_onedrive(active_bytes,  "Meeting_Issue_Log.xlsx")
                    history_link = upload_to_onedrive(history_bytes, "Meeting_Issue_History_Backup.xlsx")
                if active_link:
                    st.success("✅ อัปโหลด OneDrive สำเร็จ!")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f"📄 [เปิด Active Issues.xlsx]({active_link})")
                    with col2:
                        st.markdown(f"📜 [เปิด History Log.xlsx]({history_link})")

                # สร้าง PDF
                pdf_bytes = build_pdf_bytes(df_active, df_history, updates)

                # ส่ง Email
                if pdf_bytes:
                    with st.spinner("📧 ส่ง Email..."):
                        subject = f"รายงานการประชุม TSP Metal Works วัน{day_name}ที่ {day_num} {month_name} {year_thai}"
                        is_real_pdf = pdf_bytes[:4] == b'%PDF'
                        attach_type = "PDF" if is_real_pdf else "HTML"
                        body = f"""
<p>เรียน ทีมงานทุกท่าน</p>
<p>ขอส่งรายงานสรุปการประชุมประจำวัน{day_name}ที่ {day_num} {month_name} {year_thai} เวลา {time_str} น.</p>
<p>งานคงค้าง: <b>{len(df_active)} รายการ</b><br>
งานที่เสร็จสิ้นแล้ว: <b>{len(df_history)} รายการ</b></p>
<p>กรุณาตรวจสอบเอกสารแนบ ({attach_type}) สำหรับรายละเอียดครับ</p>
<p>ขอบคุณครับ<br>ระบบติดตามการประชุม TSP Metal Works</p>
"""
                        ok, sent_as_pdf = send_email_with_attachment(pdf_bytes, subject, body)
                        if ok:
                            file_label = "PDF" if sent_as_pdf else "HTML (WeasyPrint ไม่พร้อมใช้งาน)"
                            st.success(f"📧 ส่ง Email เรียบร้อยแล้วครับ! (ไฟล์แนบ: {file_label})")
                            if not sent_as_pdf:
                                st.info("💡 ติดตั้ง WeasyPrint บน server เพื่อส่งเป็น PDF จริงครับ: `pip install weasyprint`")

                # ส่ง LINE
                line_msg = f"🔔 อัปเดตการประชุม TSP Metal Works\nวัน{day_name}ที่ {day_num} {month_name} {year_thai} เวลา {time_str} น.\n\n"
                line_msg += f"📌 งานคงค้าง: {len(df_active)} รายการ\n"
                if not df_active.empty:
                    for _, row in df_active.iterrows():
                        line_msg += f"• {row['Issue ID']}: {row['หัวข้อปัญหา']} ({row['ผู้รับผิดชอบ']})\n"
                line_msg += f"\n✅ งานปิดแล้ว: {len(df_history)} รายการ"

                targets = [t for t in [line_user_id, line_group_id] if t]
                for t in targets:
                    send_line_message(line_msg, t)
                if targets:
                    st.success("📲 ส่ง LINE เรียบร้อยแล้วครับ!")

                # Download สำรอง
                st.divider()
                st.subheader("⬇️ ดาวน์โหลดไฟล์ Excel (สำรอง)")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.download_button("📥 Active Issues.xlsx", active_bytes, "Meeting_Issue_Log.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col2:
                    st.download_button("📥 History Log.xlsx", history_bytes, "Meeting_Issue_History_Backup.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with col3:
                    if pdf_bytes:
                        is_pdf = pdf_bytes[:4] == b'%PDF' or pdf_bytes[:5] == b'<!DOC'
                        if pdf_bytes[:5] == b'<!DOC':
                            st.download_button("📥 รายงาน HTML", pdf_bytes,
                                f"meeting_report_{now.strftime('%Y%m%d')}.html", "text/html")
                        else:
                            st.download_button("📥 รายงาน PDF", pdf_bytes,
                                f"meeting_report_{now.strftime('%Y%m%d')}.pdf", "application/pdf")

                st.cache_data.clear()

            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
