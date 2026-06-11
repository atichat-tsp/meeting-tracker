import streamlit as st
import json
import requests
import pandas as pd
import anthropic
import os
import io
import time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# =============================================
# 1. ตั้งค่าหน้าตาเว็บ
# =============================================
st.set_page_config(page_title="AI Meeting Tracker", layout="wide")
st.title("📝 ระบบบันทึกการประชุม → แจ้งเตือนสั้นผ่าน LINE & อัปโหลด Google Drive")
st.write("เวอร์ชันใช้ Claude AI | อัปโหลดไฟล์ Excel ขึ้น Google Drive อัตโนมัติ")

# =============================================
# 2. โหลด Config จาก Streamlit Secrets
# =============================================
try:
    line_bot_token   = st.secrets["LINE_BOT_TOKEN"]
    line_user_id     = st.secrets["LINE_USER_ID"]
    anthropic_key    = st.secrets["ANTHROPIC_API_KEY"]
    google_drive_url = st.secrets.get("GOOGLE_DRIVE_URL", "https://drive.google.com")
    drive_folder_id  = st.secrets["GOOGLE_DRIVE_FOLDER_ID"]
    gcp_credentials  = dict(st.secrets["gcp_service_account"])
except Exception:
    line_bot_token   = os.environ.get("LINE_BOT_TOKEN", "")
    line_user_id     = os.environ.get("LINE_USER_ID", "")
    anthropic_key    = os.environ.get("ANTHROPIC_API_KEY", "")
    google_drive_url = os.environ.get("GOOGLE_DRIVE_URL", "https://drive.google.com")
    drive_folder_id  = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
    gcp_credentials  = {}

# =============================================
# 3. ฟังก์ชัน Upload ไฟล์ขึ้น Google Drive
# =============================================
def upload_to_drive(file_bytes: bytes, filename: str, folder_id: str) -> str:
    """Upload ไฟล์ขึ้น Google Drive แล้วคืน URL"""
    try:
        creds = service_account.Credentials.from_service_account_info(
            gcp_credentials,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)

        # ลบไฟล์เก่าที่ชื่อเดียวกันออกก่อน (optional)
        results = service.files().list(
            q=f"name='{filename}' and '{folder_id}' in parents and trashed=false",
            fields="files(id, name)"
        ).execute()
        for f in results.get("files", []):
            service.files().delete(fileId=f["id"]).execute()

        # Upload ไฟล์ใหม่
        file_metadata = {"name": filename, "parents": [folder_id]}
        media = MediaIoBaseUpload(
            io.BytesIO(file_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink"
        ).execute()
        return uploaded.get("webViewLink", "")
    except Exception as e:
        st.error(f"❌ Upload Drive ไม่สำเร็จ: {e}")
        return ""

# =============================================
# 4. ฟังก์ชันส่ง LINE Notification
# =============================================
def send_line_notification(df_active, active_link="", history_link=""):
    if not line_bot_token or not line_user_id:
        st.warning("⚠️ ยังไม่ได้ตั้งค่า LINE Token — ข้ามการแจ้งเตือน")
        return

    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {line_bot_token}"
    }

    if df_active.empty:
        message_text = "🔔 Issue Tracker อัปเดต\n\nยินดีด้วยครับ! ไม่มีงานค้างแล้ว 🎉"
    else:
        lines = ["🔔 Issue Tracker อัปเดต", "สรุปหัวข้อติดตามงาน:\n"]
        for _, row in df_active.iterrows():
            lines.append(f"📋 {row['Issue ID']}: {row['หัวข้อปัญหา']}")
            lines.append(f"👤 ผู้รับผิดชอบ: {row['ผู้รับผิดชอบ']}")
            lines.append(f"📌 สถานะ: {row['สถานะ']}")
            lines.append(f"📅 กำหนดเสร็จ: {row['กำหนดเสร็จ']}")
            note = str(row.get('งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม', ''))
            if note and note.lower() not in ('none', 'nan', ''):
                lines.append(f"⚠️ บันทึก: {note}")
            lines.append("-------------------------")
        message_text = "\n".join(lines)

    if active_link:
        message_text += f"\n\n📄 Active Issues:\n{active_link}"
    if history_link:
        message_text += f"\n📜 History Log:\n{history_link}"

    payload = {"to": line_user_id, "messages": [{"type": "text", "text": message_text}]}
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=10)
        if resp.status_code == 200:
            st.success("📲 ส่งแจ้งเตือนเข้า LINE เรียบร้อยแล้วครับ!")
        else:
            st.error(f"❌ LINE ส่งไม่สำเร็จ: {resp.status_code} — {resp.text}")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดการส่ง LINE: {e}")

# =============================================
# 5. ฟังก์ชันสร้าง Excel ใน Memory
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
# 6. Prompt Template
# =============================================
PROMPT_TEMPLATE = """
คุณคือผู้ช่วยเลขานุการระดับมืออาชีพ จงวิเคราะห์บันทึกการประชุมต่อไปนี้
แล้วแยกรายการกิจกรรม (Action Items) ออกมาในรูปแบบ JSON Array
โดยแต่ละรายการต้องมีคอลัมน์ครบดังนี้ (ตอบเป็น JSON เท่านั้น):

1. "Issue ID"         : รันเลขเริ่มต้น ISS-001, ISS-002 ...
2. "หัวข้อปัญหา"       : อธิบายปัญหาหรืองานสั้นๆ กระชับชัดเจน
3. "ผู้แจ้ง"           : ชื่อผู้แจ้ง/ผู้สั่งงานจากบทสนทนา
4. "Priority"         : High / Medium / Low
5. "ผู้รับผิดชอบ"      : ชื่อผู้รับผิดชอบจากบทสนทนา
6. "ผู้ติดตาม"         : ใส่ "Teetat" เสมอ
7. "สถานะ"            : ถ้าบทสนทนาบอกว่าเสร็จ/ปิดงาน → "Closed" หรือ "Done"
                        นอกนั้น → "In Progress" หรือ "Open"
8. "งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม" : สรุปสั้นๆ
9. "ความคิดเห็น / Comments" : ข้อความสรุปพูดคุยสั้นๆ
10. "กำหนดเสร็จ"      : รูปแบบ DD/MM/YYYY คำนวณจากบริบทปี 2026

บันทึกการประชุม:
\"\"\"
{raw_notes}
\"\"\"
"""

EXPECTED_COLUMNS = [
    '#', 'Issue ID', 'หัวข้อปัญหา', 'ผู้แจ้ง', 'Priority',
    'ผู้รับผิดชอบ', 'ผู้ติดตาม', 'สถานะ',
    'งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม',
    'ความคิดเห็น / Comments', 'กำหนดเสร็จ'
]

# =============================================
# 7. UI
# =============================================
default_text = """บอส: สัปดาห์หน้าเราต้องเคลียร์เรื่องกล้องวงจรปิดในไลน์ผลิตใหม่ให้เสร็จนะ
พี่เกรียงศักดิ์: ได้ครับ เดี๋ยวผมไปเคลียร์หน้างานและหาซัพพลายเออร์เข้ามาประเมินราคาให้เสร็จภายในวันศุกร์หน้าที่ 19 ครับ แต่ตอนนี้แอบติดปัญหาเรื่องแบบแปลนอาคารเก่านิดหน่อย ขาดข้อมูลโยธาครับ
บอส: ปฏิพาท เรื่องเซ็ต VLAN ไปถึงไหนแล้ว?
ปฏิพาท: อันนี้ผมคอนฟิกและเดินสายเสร็จหมดแล้วครับ ทดสอบระบบเรียบร้อย ปิดงานได้เลยครับบอส
บอส: ดีมาก สรุปตามนี้"""

raw_notes = st.text_area(
    "✍️ พิมพ์บันทึกการประชุมดิบที่นี่:",
    value=default_text,
    height=220,
    placeholder="วางบันทึกประชุมดิบได้เลยครับ ไม่จำกัดความยาว..."
)

# =============================================
# 8. ปุ่มประมวลผล
# =============================================
if st.button("🚀 ประมวลผล + อัปโหลด Google Drive + แจ้งเตือน LINE", type="primary"):
    if not raw_notes.strip():
        st.warning("⚠️ กรุณาพิมพ์บันทึกการประชุมก่อนครับ")
    elif not anthropic_key:
        st.error("❌ ไม่พบ ANTHROPIC_API_KEY")
    else:
        with st.spinner("🤖 Claude AI กำลังวิเคราะห์บันทึกการประชุม..."):
            try:
                client = anthropic.Anthropic(api_key=anthropic_key)

                max_retries = 3
                response = None
                for attempt in range(max_retries):
                    try:
                        response = client.messages.create(
                            model="claude-sonnet-4-6",
                            max_tokens=8192,
                            messages=[{
                                "role": "user",
                                "content": PROMPT_TEMPLATE.format(raw_notes=raw_notes)
                            }]
                        )
                        break
                    except anthropic.APIStatusError as e:
                        if attempt < max_retries - 1:
                            st.warning(f"⏳ พยายามอีกครั้ง ({attempt + 2}/{max_retries})...")
                            time.sleep(3)
                        else:
                            raise e

                raw_text = response.content[0].text.strip()
                if raw_text.startswith("```"):
                    raw_text = raw_text.split("```")[1]
                    if raw_text.startswith("json"):
                        raw_text = raw_text[4:]
                raw_text = raw_text.strip()

                data_json = json.loads(raw_text)
                df_all = pd.DataFrame(data_json)
                df_all.insert(0, '#', range(1, len(df_all) + 1))
                df_all.columns = EXPECTED_COLUMNS

                done_statuses = {'closed', 'done'}
                mask_done = df_all['สถานะ'].str.lower().str.strip().isin(done_statuses)
                df_active  = df_all[~mask_done].copy().reset_index(drop=True)
                df_history = df_all[ mask_done].copy().reset_index(drop=True)

                active_bytes  = build_excel_bytes(df_active,  'Active Issues')
                history_bytes = build_excel_bytes(df_history, 'History Log')

                # แสดงตาราง
                st.subheader(f"📌 งานคงค้าง (Active) — {len(df_active)} รายการ")
                if df_active.empty:
                    st.info("ไม่มีงานค้างครับ 🎉")
                else:
                    st.dataframe(df_active, use_container_width=True)

                st.subheader(f"📜 งานที่สิ้นสุดแล้ว (History) — {len(df_history)} รายการ")
                if df_history.empty:
                    st.info("ไม่มีงานที่ปิดแล้วในรอบนี้ครับ")
                else:
                    st.dataframe(df_history, use_container_width=True)

                # Upload Google Drive
                st.divider()
                active_link  = ""
                history_link = ""

                if drive_folder_id and gcp_credentials:
                    with st.spinner("☁️ กำลังอัปโหลดขึ้น Google Drive..."):
                        active_link  = upload_to_drive(active_bytes,  "Meeting_Issue_Log.xlsx",            drive_folder_id)
                        history_link = upload_to_drive(history_bytes, "Meeting_Issue_History_Backup.xlsx", drive_folder_id)

                    if active_link:
                        st.success("✅ อัปโหลด Google Drive สำเร็จแล้วครับ!")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown(f"📄 [เปิด Active Issues.xlsx]({active_link})")
                        with col2:
                            st.markdown(f"📜 [เปิด History Log.xlsx]({history_link})")

                # ปุ่ม Download สำรอง
                st.divider()
                st.subheader("⬇️ ดาวน์โหลดไฟล์ Excel (สำรอง)")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 ดาวน์โหลด Active Issues.xlsx",
                        data=active_bytes,
                        file_name="Meeting_Issue_Log.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.download_button(
                        label="📥 ดาวน์โหลด History Log.xlsx",
                        data=history_bytes,
                        file_name="Meeting_Issue_History_Backup.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # ส่ง LINE
                send_line_notification(df_active, active_link, history_link)

            except json.JSONDecodeError as e:
                st.error(f"❌ Claude ตอบกลับไม่ใช่ JSON: {e}")
            except anthropic.AuthenticationError:
                st.error("❌ API Key ไม่ถูกต้อง")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
