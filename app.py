import streamlit as st
import json
import requests
import pandas as pd
import anthropic
import os
import io
import time

# =============================================
# 1. ตั้งค่าหน้าตาเว็บ
# =============================================
st.set_page_config(page_title="AI Meeting Tracker", layout="wide")
st.title("📝 ระบบบันทึกการประชุม → แจ้งเตือนสั้นผ่าน LINE & แยกไฟล์ประวัติ")
st.write("เวอร์ชันใช้ Claude AI | เพิ่มคอลัมน์อุปสรรค/งานที่ทำ และแยกเก็บไฟล์งานที่จบแล้ว (Done/Closed) อัตโนมัติ")

# =============================================
# 2. โหลด Config จาก Streamlit Secrets
# =============================================
try:
    line_bot_token = st.secrets["LINE_BOT_TOKEN"]
    line_user_id   = st.secrets["LINE_USER_ID"]
    anthropic_key  = st.secrets["ANTHROPIC_API_KEY"]
    google_drive_url = st.secrets.get("GOOGLE_DRIVE_URL", "https://drive.google.com")
except Exception:
    # fallback สำหรับ dev local (ใช้ .env หรือใส่ตรงนี้ชั่วคราว)
    line_bot_token   = os.environ.get("LINE_BOT_TOKEN", "")
    line_user_id     = os.environ.get("LINE_USER_ID", "")
    anthropic_key    = os.environ.get("ANTHROPIC_API_KEY", "")
    google_drive_url = os.environ.get("GOOGLE_DRIVE_URL", "https://drive.google.com")

# =============================================
# 3. ฟังก์ชันส่ง LINE Notification
# =============================================
def send_line_notification(df_active):
    if not line_bot_token or not line_user_id:
        st.warning("⚠️ ยังไม่ได้ตั้งค่า LINE Token — ข้ามการแจ้งเตือน")
        return

    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {line_bot_token}"
    }

    if df_active.empty:
        message_text = (
            "🔔 Issue Tracker อัปเดต\n\n"
            "ยินดีด้วยครับ! ตอนนี้ไม่มีหัวข้องานค้างสำหรับการประชุมครั้งถัดไปแล้วครับ 🎉"
        )
    else:
        lines = ["🔔 Issue Tracker อัปเดต", "สรุปหัวข้อติดตามงานในการประชุมครั้งถัดไป:\n"]
        for _, row in df_active.iterrows():
            lines.append(f"📋 {row['Issue ID']}: {row['หัวข้อปัญหา']}")
            lines.append(f"👤 ผู้รับผิดชอบ: {row['ผู้รับผิดชอบ']}")
            lines.append(f"📌 สถานะ: {row['สถานะ']}")
            lines.append(f"📅 กำหนดเสร็จ: {row['กำหนดเสร็จ']}")
            note = str(row.get('งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม', ''))
            if note and note.lower() not in ('none', 'nan', ''):
                lines.append(f"⚠️ บันทึกหน้างาน: {note}")
            lines.append("-------------------------")
        message_text = "\n".join(lines)

    message_text += f"\n\n📁 Google Drive:\n{google_drive_url}"

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
# 4. ฟังก์ชันสร้าง Excel ใน Memory (BytesIO)
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
# 5. Prompt Template
# =============================================
PROMPT_TEMPLATE = """
คุณคือผู้ช่วยเลขานุการระดับมืออาชีพ จงวิเคราะห์บันทึกการประชุมต่อไปนี้
แล้วแยกรายการกิจกรรม (Action Items) ออกมาในรูปแบบ JSON Array
โดยแต่ละรายการต้องมีคอลัมน์ครบดังนี้ (ตอบเป็น JSON เท่านั้น ห้ามมีข้อความอื่นนำหน้าหรือต่อท้าย):

1. "Issue ID"         : รันเลขเริ่มต้น ISS-001, ISS-002 ...
2. "หัวข้อปัญหา"       : อธิบายปัญหาหรืองานสั้นๆ กระชับชัดเจน
3. "ผู้แจ้ง"           : ชื่อผู้แจ้ง/ผู้สั่งงานจากบทสนทนา
4. "Priority"         : High / Medium / Low
5. "ผู้รับผิดชอบ"      : ชื่อผู้รับผิดชอบจากบทสนทนา
6. "ผู้ติดตาม"         : ใส่ "Teetat" เสมอ
7. "สถานะ"            : ถ้าบทสนทนาบอกว่าเสร็จ/ปิดงาน → "Closed" หรือ "Done"
                        นอกนั้น → "In Progress" หรือ "Open"
8. "งานที่ทำไปแล้ว / อุปสรรค / Resource ที่ต้องการเพิ่ม" :
   สรุปสั้นๆ ว่าทำอะไรไปแล้ว ติดปัญหาอะไร หรือต้องรออะไรเพิ่ม
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
# 6. UI — กล่องรับข้อความ
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
# 7. ปุ่มประมวลผล
# =============================================
if st.button("🚀 ประมวลผลและแยกไฟล์ Excel & แจ้งเตือน LINE", type="primary"):
    if not raw_notes.strip():
        st.warning("⚠️ กรุณาพิมพ์บันทึกการประชุมก่อนครับ")
    elif not anthropic_key:
        st.error("❌ ไม่พบ ANTHROPIC_API_KEY — กรุณาตั้งค่าใน Streamlit Secrets ครับ")
    else:
        with st.spinner("🤖 Claude AI กำลังวิเคราะห์บันทึกการประชุม..."):
            try:
                client = anthropic.Anthropic(api_key=anthropic_key)

                # เรียก Claude API พร้อม retry อัตโนมัติ
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

                # แยก JSON ออกจาก response
                raw_text = response.content[0].text.strip()

                # ลอง clean JSON กรณี Claude ใส่ backtick มาด้วย
                if raw_text.startswith("```"):
                    raw_text = raw_text.split("```")[1]
                    if raw_text.startswith("json"):
                        raw_text = raw_text[4:]
                raw_text = raw_text.strip()

                data_json = json.loads(raw_text)
                df_all = pd.DataFrame(data_json)

                # จัดคอลัมน์ให้ตรง
                df_all.insert(0, '#', range(1, len(df_all) + 1))
                df_all.columns = EXPECTED_COLUMNS

                # แยก Active vs History
                done_statuses = {'closed', 'done'}
                mask_done = df_all['สถานะ'].str.lower().str.strip().isin(done_statuses)
                df_active  = df_all[~mask_done].copy().reset_index(drop=True)
                df_history = df_all[ mask_done].copy().reset_index(drop=True)

                # สร้างไฟล์ Excel ใน memory
                active_bytes  = build_excel_bytes(df_active,  'Active Issues')
                history_bytes = build_excel_bytes(df_history, 'History Log')

                # =============================================
                # แสดงผลตาราง
                # =============================================
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

                # =============================================
                # ปุ่ม Download Excel
                # =============================================
                st.divider()
                st.subheader("⬇️ ดาวน์โหลดไฟล์ Excel")

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

                st.markdown(f"📁 **[เปิด Google Drive เพื่อนำไฟล์ไปอัปโหลด]({google_drive_url})**")

                # =============================================
                # ส่ง LINE (เฉพาะ Active)
                # =============================================
                send_line_notification(df_active)

            except json.JSONDecodeError as e:
                st.error(f"❌ Claude ตอบกลับไม่ใช่ JSON ที่ถูกต้อง: {e}")
                st.code(raw_text[:2000], language="text")
            except anthropic.AuthenticationError:
                st.error("❌ API Key ไม่ถูกต้อง — กรุณาตรวจสอบ ANTHROPIC_API_KEY")
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาด: {e}")
