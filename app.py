import streamlit as str_app
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import os
import json
import random
import time
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

str_app.set_page_config(
    page_title="⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",  # <--- Thêm đúng dòng này vào đây
)

PROGRESS_FILE = "quiz_progress.json"

SENDER_EMAIL = "ducanhdao74@gmail.com"
SENDER_PASSWORD = "ospeifebafqlufpi"

def load_saved_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {
    "completed_chunks": [0, 1, 2],
    "passed_tests": [0, 1, 2],
    "bookmarked_questions": [],
    "wrong_questions": [],
    "spaced_repetition_data": {},
    "total_study_seconds": 0,
    "last_login_date": ""
}
import subprocess

def save_current_progress_and_sync_github():
    # 1. Ghi dữ liệu trực tiếp vào file JSON cục bộ
    data = {
        "completed_chunks": list(str_app.session_state.get("completed_chunks", [])),
        "passed_tests": list(str_app.session_state.get("passed_tests", [])),
        "bookmarked_questions": str_app.session_state.get("bookmarked_questions", []),
        "wrong_questions": str_app.session_state.get("wrong_questions", []),
        "spaced_repetition_data": str_app.session_state.get("spaced_repetition_data", {}),
        "total_study_seconds": str_app.session_state.get("total_study_seconds", 0),
        "last_login_date": str_app.session_state.get("last_login_date", "")
    }
    
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        # 2. Tự động Git commit & push lên GitHub
        subprocess.run(["git", "config", "--global", "user.email", "ducanh@bot.com"], check=False)
        subprocess.run(["git", "config", "--global", "user.name", "Quiz Bot Auto Sync"], check=False)
        subprocess.run(["git", "add", PROGRESS_FILE], check=True)
        commit_result = subprocess.run(["git", "commit", "-m", "Auto-update quiz progress json"], capture_output=True, text=True)
        
        if "nothing to commit" not in commit_result.stdout:
            subprocess.run(["git", "push"], check=True)
        return True
    except Exception as e:
        print(f"Lỗi đồng bộ Git tự động: {e}")
        return False

def save_current_progress():
    # Hàm cầu nối đặt ở phía sau hàm chính để hứng lệnh gọi cũ an toàn
    save_current_progress_and_sync_github()
def scroll_to_top():
    components.html("""
        <script>
            const doc = window.parent.document;
            setTimeout(() => {
                const main = doc.querySelector('.main') || doc.documentElement;
                main.scrollTo({top: 0, behavior: 'auto'});
                window.parent.scrollTo({top: 0, behavior: 'auto'});
            }, 10);
        </script>
    """, height=0)

def send_daily_reminder_email(
    receiver_email,
    completed_questions_count,
    total_questions_count,
    bookmarked_count,
    total_study_hours,
    is_completion=False,
):
  if total_study_hours < 1:
    time_str = f"{int(total_study_hours * 60)} phút"
  else:
    time_str = f"{total_study_hours:.1f} giờ"

  if is_completion:
    subject = "Bao cao tong ket hoan thanh phien on tap An Toan Dien!"
    body = f"""Chao Duc Anh,

He thong ghi nhan ong vua hoan thanh mot phien on tap chuan bi thi Dien luc Dien Bien:
- Thoi gian ket thuc phien: Hom nay
- Trang thai: Da hoan thanh on luyen va reset phien lam viec ve trang chu.

Tien do phien vua roi:
- So cau da hoan thanh: {completed_questions_count}/{total_questions_count} cau
- Tong thoi gian on tap: {time_str}
- So cau hoi can luu y (Star): {bookmarked_count} cau

Chuc ong tiep tuc giu vung phong do cho ky thi sap toi!"""
  else:
    subject = "Nhac nho on tap An Toan Dien moi ngay!"
    body = f"""Chao Duc Anh,

Hom nay la mot ngay moi roi! Hay danh ra chut thoi gian de vao on tap ngan hang cau hoi An Toan Dien nhe:

Tien do hien tai cua ong:
- So cau da hoan thanh: {completed_questions_count}/{total_questions_count} cau
- Tong thoi gian da on tap: {time_str}
- So cau hoi dang can ghi nho (Star): {bookmarked_count} cau

Chuc ong on thi that tot va dat ket qua cao!"""

  message = MIMEMultipart()
  message["From"] = SENDER_EMAIL
  message["To"] = receiver_email
  message["Subject"] = Header(subject, "utf-8")
  message.attach(MIMEText(body, "plain", "utf-8"))

  try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)
    server.sendmail(SENDER_EMAIL, receiver_email, message.as_string())
    server.quit()
    return True, "Thành công"
  except Exception as e:
    return False, str(e)

@str_app.cache_data
def load_data():
    file_name = "PL1. Tong hop ngan hang cau hoi an toàn nam 2025 fn (1).xlsx"
    if not os.path.exists(file_name):
        for f in os.listdir("."):
            if f.endswith(".xlsx") and "ngan hang cau hoi" in f.lower():
                file_name = f
                break
    if not os.path.exists(file_name):
        return None, f"Không tìm thấy file Excel trong thư mục!"
    try:
        wb = openpyxl.load_workbook(file_name, data_only=True)
        sheets_data = {}
        
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            questions = []
            current_q = None
            current_opts = []
            correct_ans = "A"
            
            for row in ws.iter_rows(values_only=False):
                cell = row[0]
                val = cell.value
                if val is not None:
                    val_str = str(val).strip()
                    is_red = False
                    if cell.font and cell.font.color:
                        c = cell.font.color
                        if c.rgb:
                            rgb_str = str(c.rgb).upper()
                            if any(x in rgb_str for x in ["FF0000", "ED1C24", "C00000", "RED"]):
                                is_red = True
                            elif len(rgb_str) >= 6:
                                try:
                                    hex_val = rgb_str[-6:]
                                    r_val = int(hex_val[0:2], 16)
                                    g_val = int(hex_val[2:4], 16)
                                    b_val = int(hex_val[4:6], 16)
                                    if r_val > 150 and g_val < 80 and b_val < 80:
                                        is_red = True
                                except:
                                    pass
                        if hasattr(c, 'indexed') and c.indexed in [10, 2]:
                            is_red = True

                    if val_str.lower().startswith("câu"):
                        if current_q:
                            questions.append({
                                "question": current_q,
                                "options": current_opts,
                                "correct": correct_ans,
                                "sheet": sheet
                            })
                        current_q = val_str
                        current_opts = []
                        correct_ans = "A"
                    elif val_str.lower().startswith(("a.", "b.", "c.", "d.")):
                        current_opts.append(val_str)
                        if is_red:
                            correct_ans = val_str[0].upper()
                    else:
                        if current_q and not current_opts:
                            current_q += " " + val_str
                        elif current_opts:
                            current_opts[-1] += " " + val_str
                            
            if current_q:
                questions.append({
                    "question": current_q,
                    "options": current_opts,
                    "correct": correct_ans,
                    "sheet": sheet
                })
                
            sheets_data[sheet] = questions
        return sheets_data, None
    except Exception as e:
        return None, f"Lỗi đọc file: {str(e)}"

if "start_session_time" not in str_app.session_state:
    str_app.session_state["start_session_time"] = time.time()

sheets_data, error_message = load_data()
saved_prog = load_saved_progress()

total_all_questions = 0
if sheets_data:
    for sh, ql in sheets_data.items():
        total_all_questions += len(ql)

if "wrong_questions" not in str_app.session_state:
    str_app.session_state["wrong_questions"] = saved_prog.get("wrong_questions", [])
if "bookmarked_questions" not in str_app.session_state:
    str_app.session_state["bookmarked_questions"] = saved_prog.get("bookmarked_questions", [])
if "spaced_repetition_data" not in str_app.session_state:
    str_app.session_state["spaced_repetition_data"] = saved_prog.get("spaced_repetition_data", {})
if "completed_chunks" not in str_app.session_state:
  # Thêm trực tiếp [0, 1, 2] vào để ép nhận 3 phần đầu đã hoàn thành
  saved_chunks = saved_prog.get("completed_chunks", [])
  if not saved_chunks:
    saved_chunks = [0, 1, 2]
  str_app.session_state["completed_chunks"] = set(saved_chunks)

if "passed_tests" not in str_app.session_state:
  saved_passed = saved_prog.get("passed_tests", [])
  if not saved_passed:
    saved_passed = [0, 1, 2]
  str_app.session_state["passed_tests"] = set(saved_passed)

save_current_progress()

def update_spaced_repetition(q_text, is_correct):
    sr_data = str_app.session_state["spaced_repetition_data"]
    now = time.time()
    if q_text not in sr_data:
        sr_data[q_text] = {"box": 1, "next_review": 0}
    
    item = sr_data[q_text]
    if is_correct:
        item["box"] = min(5, item["box"] + 1)
        intervals = [0, 0, 4*3600, 24*3600, 3*24*3600, 7*24*3600]
        item["next_review"] = now + intervals[item["box"]]
    else:
        item["box"] = 1
        item["next_review"] = now
    save_current_progress()

def scheduled_job():
    saved_data = load_saved_progress()
    bm_cnt = len(saved_data.get("bookmarked_questions", []))
    completed_chunks_set = set(saved_data.get("completed_chunks", [])).union(set(saved_data.get("passed_tests", [])))
    comp_q_cnt = min(len(completed_chunks_set) * 50, total_all_questions)
    total_hours = saved_data.get("total_study_seconds", 0) / 3600.0
    send_daily_reminder_email(SENDER_EMAIL, comp_q_cnt, total_all_questions, bm_cnt, total_hours)

@str_app.cache_resource
def start_scheduler():
    scheduler = BackgroundScheduler()
    scheduler.add_job(scheduled_job, 'cron', hour=7, minute=0)
    scheduler.start()
    return scheduler

try:
    start_scheduler()
except Exception:
    pass

# --- CYBERPUNK / DARK TECH GLASSMORPHISM CSS & OPTION CARDS ---
str_app.markdown("""
    <style>
    @keyframes shine {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .sparkle-title {
        background: linear-gradient(270deg, #38bdf8, #818cf8, #c084fc, #38bdf8);
        background-size: 300% 300%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: shine 6s ease infinite;
        font-weight: 800;
        font-size: 2.8rem;
    }
    .welcome-card {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.85), rgba(30, 41, 59, 0.75));
        border: 1px solid rgba(56, 189, 248, 0.35);
        box-shadow: 0 0 30px rgba(56, 189, 248, 0.2);
        border-radius: 20px;
        padding: 40px;
        text-align: center;
        backdrop-filter: blur(12px);
    }
    .metric-card-container {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.7));
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 14px;
        padding: 14px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        margin-bottom: 12px;
        backdrop-filter: blur(8px);
        transition: all 0.3s ease;
    }
    .metric-card-container:hover {
        border-color: rgba(56, 189, 248, 0.7);
        box-shadow: 0 0 25px rgba(56, 189, 248, 0.35);
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: 800;
        color: #38bdf8;
        margin-top: 2px;
        text-shadow: 0 0 10px rgba(56, 189, 248, 0.4);
    }
    .metric-label {
        font-size: 0.8rem;
        color: #94a3b8;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .main-header-card {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.6));
        border-left: 5px solid #38bdf8;
        border-top: 1px solid rgba(56, 189, 248, 0.2);
        border-right: 1px solid rgba(56, 189, 248, 0.2);
        border-bottom: 1px solid rgba(56, 189, 248, 0.2);
        padding: 16px 20px;
        border-radius: 0 12px 12px 0;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        backdrop-filter: blur(8px);
    }
    .question-card {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.85), rgba(30, 41, 59, 0.75));
        border: 1px solid rgba(56, 189, 248, 0.25);
        border-left: 6px solid #38bdf8;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        margin-bottom: 20px;
        backdrop-filter: blur(8px);
    }
    .question-title {
        font-size: 1.15rem !important;
        line-height: 1.6;
        color: #f8fafc !important;
        font-weight: 600;
        margin: 0;
    }
    
    /* --- NÂNG CẤP GIAO DIỆN CHỌN ĐÁP ÁN (OPTION CARDS) --- */
    div[data-testid="stRadio"] > div[role="radiogroup"] {
        gap: 12px !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.8), rgba(30, 41, 59, 0.7)) !important;
        border: 1px solid rgba(56, 189, 248, 0.25) !important;
        border-radius: 12px !important;
        padding: 14px 18px !important;
        width: 100% !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2) !important;
        transition: all 0.3s ease !important;
        cursor: pointer !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label:hover {
        border-color: #38bdf8 !important;
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.95), rgba(56, 189, 248, 0.2)) !important;
        box-shadow: 0 0 20px rgba(56, 189, 248, 0.35) !important;
        transform: translateX(4px);
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label p,
    div[data-testid="stRadio"] > div[role="radiogroup"] label span,
    div[data-testid="stRadio"] > div[role="radiogroup"] label div {
        font-size: 1.05rem !important;
        line-height: 1.5 !important;
        color: #f1f5f9 !important;
        font-weight: 500 !important;
    }
    </style>
""", unsafe_allow_html=True)

def get_shuffled_options(q_item, session_key):
    if session_key not in str_app.session_state:
        orig_options = q_item["options"]
        if not orig_options:
            orig_options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
        
        correct_letter = q_item.get("correct", "A").strip().upper()
        
        opts_with_status = []
        for opt in orig_options:
            opt_prefix = opt[:1].strip().upper()
            is_corr = (opt_prefix == correct_letter)
            clean_text = opt[2:].strip() if len(opt) > 2 and opt[1] in [".", ")"] else opt
            opts_with_status.append({"text": clean_text, "is_correct": is_corr})
            
        random.shuffle(opts_with_status)
        
        labels = ["A", "B", "C", "D"]
        shuffled_options = []
        new_correct_letter = "A"
        
        for i, item in enumerate(opts_with_status):
            lbl = labels[i] if i < len(labels) else str(i + 1)
            formatted_opt = f"{lbl}. {item['text']}"
            shuffled_options.append(formatted_opt)
            if item["is_correct"]:
                new_correct_letter = lbl
                
        str_app.session_state[session_key] = {
            "options": shuffled_options,
            "correct": new_correct_letter
        }
        
str_app.sidebar.markdown(
      f"""
<div class="metric-card-container">
<div class="metric-label">📈 Tiến độ hoàn thành</div>
<div class="metric-value">{progress_ratio * 100:.1f}%</div>
<div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">{completed_est_count} / {total_all_questions} câu</div>
</div>
""",
      unsafe_allow_html=True,
  )
  # 1. Tính toán các biến số trước
completed_chunks_set = str_app.session_state["completed_chunks"].union(
      str_app.session_state["passed_tests"]
  )
chunk_size_calc = 50
completed_est_count = min(
      len(completed_chunks_set) * chunk_size_calc, total_all_questions
  )
progress_ratio = (
      completed_est_count / total_all_questions if total_all_questions > 0 else 0.0
  )

current_total_seconds = saved_prog.get("total_study_seconds", 0) + (
time.time() - str_app.session_state["start_session_time"]
  )
bookmarked_count = len(str_app.session_state.get("bookmarked questions", []))
wrong_count = len(str_app.session_state.get("wrong_questions", []))
flagged_count = len(str_app.session_state.get("flagged_questions", []))

  # 2. Hiển thị Dashboard Sidebar
str_app.sidebar.markdown(
      """
        <div style="font-size: 1.2rem; font-weight: 800; color: #f8fafc; margin-bottom: 15px; display: flex; align-items: center; gap: 8px;">
            ⚡ Dashboard Tổng Quan
        </div>
    """,
      unsafe_allow_html=True,
  )
  str_app.sidebar.markdown(
      f"""
        <div class="metric-card-container">
            <div class="metric-label">📈 Tiến độ hoàn thành</div>
            <div class="metric-value">{progress_ratio * 100:.1f}%</div>
            <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">{completed_est_count}/{total_all_questions} câu</div>
        </div>
    """,
      unsafe_allow_html=True,
  )
  str_app.sidebar.progress(progress_ratio)

if str_app.session_state.get("is_studying", False):
  str_app.sidebar.markdown(
      f"""
        <div class="metric-card-container">
            <div class="metric-label">📈 Tiến độ hoàn thành</div>
            <div class="metric-value">{progress_ratio * 100:.1f}%</div>
            <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 4px;">{completed_est_count}/{total_all_questions} câu</div>
        </div>
    """,
      unsafe_allow_html=True,
  )
  str_app.sidebar.progress(progress_ratio)

  col_s1, col_s2 = str_app.sidebar.columns(2)
  with col_s1:
    str_app.sidebar.markdown(
        f"""
        <div class="metric-card-container" style="padding: 10px 6px;">
            <div class="metric-label" style="font-size: 0.7rem;">⭐ Ghi nhớ</div>
            <div class="metric-value" style="font-size: 1.2rem; color: #eab308; text-shadow: 0 0 10px rgba(234, 179, 8, 0.4);">{bookmarked_count}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
  with col_s2:
    str_app.sidebar.markdown(
        f"""
        <div class="metric-card-container" style="padding: 10px 6px;">
            <div class="metric-label" style="font-size: 0.7rem;">❌ Sai nhiều</div>
            <div class="metric-value" style="font-size: 1.2rem; color: #ef4444; text-shadow: 0 0 10px rgba(239, 68, 68, 0.4);">{wrong_count}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    str_app.sidebar.markdown(f"""
        <div class="metric-card-container">
            <div class="metric-label">⏱️ Tổng thời gian ôn</div>
            <div class="metric-value" style="font-size: 1.4rem; color: #818cf8; text-shadow: 0 0 10px rgba(129, 140, 248, 0.4);">{current_total_hours:.2f}h</div>
            <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 2px;">~{int(current_total_seconds // 60)} phút tập trung</div>
        </div>
    """, unsafe_allow_html=True)

    str_app.sidebar.markdown("<br>", unsafe_allow_html=True)

    if str_app.sidebar.button(
    "💾 Lưu lại tiến độ học", type="primary", use_container_width=True
):
        save_current_progress_and_sync_github()
        str_app.sidebar.success("✅ Đã lưu và đồng bộ tiến độ lên GitHub!")

    if str_app.sidebar.button("📧 Gửi Email nhắc nhở báo cáo", use_container_width=True):
        success, err_msg = send_daily_reminder_email(SENDER_EMAIL, completed_est_count, total_all_questions, bookmarked_count, current_total_hours)
        if success:
            str_app.sidebar.success("✅ Đã gửi email báo cáo vào Gmail!")
        else:
            str_app.sidebar.error(f"❌ Gửi mail thất bại: {err_msg}")

    str_app.sidebar.markdown("<br>", unsafe_allow_html=True)
    
    str_app.sidebar.markdown("""
        <div class="metric-card-container" style="text-align: left;">
            <div class="metric-label" style="margin-bottom: 8px;">🎛️ Điều Hướng Học Tập</div>
    """, unsafe_allow_html=True)
    
    mode = str_app.sidebar.selectbox("Chọn chế độ học:", [
        "📖 Ôn tập theo chuyên đề",
        "🔄 Ôn lại câu trả lời sai",
        "🧠 Spaced Repetition (Ôn thông minh)",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "⭐ Tất cả câu hỏi cần ghi nhớ",
        "📝 Thi thử (Mock Test)",
        "⚙️ Quản lý kho lưu trữ & Dữ liệu"
    ], label_visibility="collapsed", key="sidebar_mode_select")
    
    str_app.sidebar.markdown("</div>", unsafe_allow_html=True)

    if error_message:
        str_app.error(error_message)
    else:
        if mode == "📖 Ôn tập theo chuyên đề":
            str_app.sidebar.markdown("""
                <div class="metric-card-container" style="text-align: left;">
                    <div class="metric-label" style="margin-bottom: 8px;">📂 Chọn Chuyên Đề</div>
            """, unsafe_allow_html=True)
            sheet_list = list(sheets_data.keys())
            selected_sheet = str_app.sidebar.selectbox("Chuyên đề:", sheet_list, label_visibility="collapsed", key="sidebar_chuande_select")
            str_app.sidebar.markdown("</div>", unsafe_allow_html=True)
            
            q_list = sheets_data[selected_sheet]
            total_q = len(q_list)
            
            if f"q_idx_{selected_sheet}" not in str_app.session_state:
                str_app.session_state[f"q_idx_{selected_sheet}"] = 0
            if f"done_{selected_sheet}" not in str_app.session_state:
                str_app.session_state[f"done_{selected_sheet}"] = 0
                
            col_sb1, col_sb2 = str_app.sidebar.columns(2)
            col_sb1.metric("Tổng câu", total_q)
            col_sb2.metric("Đã làm", str_app.session_state[f'done_{selected_sheet}'])
            
            if str_app.sidebar.button("🔄 Đặt lại chuyên đề này", use_container_width=True):
                str_app.session_state[f"q_idx_{selected_sheet}"] = 0
                str_app.session_state[f"done_{selected_sheet}"] = 0
                keys_to_del = [k for k in str_app.session_state.keys() if k.startswith(f"shuff_chuande_{selected_sheet}_") or k.startswith(f"user_ans_chuande_{selected_sheet}_") or k.startswith(f"answered_chuande_{selected_sheet}_")]
                for k in keys_to_del:
                    del str_app.session_state[k]
                scroll_to_top()
                str_app.rerun()

            str_app.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
            str_app.markdown("---")
            
            idx = str_app.session_state[f"q_idx_{selected_sheet}"]
            if idx >= total_q and total_q > 0:
                idx = 0
                str_app.session_state[f"q_idx_{selected_sheet}"] = 0
                
            if total_q == 0:
                str_app.warning("Chuyên đề này hiện chưa có câu hỏi nào.")
            else:
                q_item = q_list[idx]
                
                str_app.markdown(f"""
                    <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">📂 Chuyên đề: {selected_sheet}</span>
                            <span class="badge-topic">Câu {idx + 1} / {total_q}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                is_bm = any(b.get("question") == q_item["question"] for b in str_app.session_state["bookmarked_questions"])
                bm_label = "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
                
                col_q_head1, col_q_head2 = str_app.columns([4, 1])
                with col_q_head2:
                    if str_app.button(bm_label, key=f"bm_chuande_{idx}", use_container_width=True):
                        if is_bm:
                            str_app.session_state["bookmarked_questions"] = [b for b in str_app.session_state["bookmarked_questions"] if b.get("question") != q_item["question"]]
                            str_app.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
                        else:
                            str_app.session_state["bookmarked_questions"].append(q_item)
                            str_app.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
                        save_current_progress()
                        str_app.rerun()
                
                clean_q_text = q_item['question']
                if clean_q_text.lower().startswith("câu"):
                    parts = clean_q_text.split(":", 1)
                    if len(parts) > 1:
                        clean_q_text = parts[1].strip()

                str_app.markdown(f"""
                    <div class="question-card">
                        <div class="question-title">
                            <b>Câu {idx + 1}:</b> {clean_q_text}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                shuff_data = get_shuffled_options(q_item, f"shuff_chuande_{selected_sheet}_{idx}")
                options = shuff_data["options"]
                correct_letter = shuff_data["correct"]
                
                ans_storage_key = f"user_ans_chuande_{selected_sheet}_{idx}"
                answered_key = f"answered_chuande_{selected_sheet}_{idx}"
                is_already_answered = str_app.session_state.get(answered_key, False)
                
                default_idx = None
                current_saved_ans = str_app.session_state.get(ans_storage_key, None)
                if current_saved_ans in options:
                    default_idx = options.index(current_saved_ans)
                    
                if not is_already_answered:
                    selected_opt = str_app.radio("Lựa chọn đáp án:", options, index=default_idx, key=f"radio_chuande_{selected_sheet}_{idx}", label_visibility="collapsed")
                    
                    if selected_opt is not None:
                        str_app.session_state[ans_storage_key] = selected_opt
                        str_app.session_state[answered_key] = True
                        str_app.session_state[f"done_{selected_sheet}"] = min(total_q, max(str_app.session_state[f"done_{selected_sheet}"], idx + 1))
                        
                        is_correct = selected_opt.strip().upper().startswith(correct_letter)
                        update_spaced_repetition(q_item["question"], is_correct)
                        if not is_correct:
                            if not any(w.get("question") == q_item["question"] for w in str_app.session_state["wrong_questions"]):
                                str_app.session_state["wrong_questions"].append(q_item)
                                save_current_progress()
                        str_app.rerun()
                else:
                    saved_choice = str_app.session_state.get(ans_storage_key)
                    str_app.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_choice}</b></p>", unsafe_allow_html=True)
                    
                    is_correct = saved_choice.strip().upper().startswith(correct_letter)
                    str_app.markdown("<br>", unsafe_allow_html=True)
                    if is_correct:
                        str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                    else:
                        correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                        str_app.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
                
                str_app.markdown("<br>", unsafe_allow_html=True)
                col_prev, col_next = str_app.columns([1, 1])
                with col_prev:
                    if str_app.button("⬅️ Câu trước", use_container_width=True):
                        if str_app.session_state[f"q_idx_{selected_sheet}"] > 0:
                            str_app.session_state[f"q_idx_{selected_sheet}"] -= 1
                        else:
                            str_app.session_state[f"q_idx_{selected_sheet}"] = total_q - 1
                        scroll_to_top()
                        str_app.rerun()
                with col_next:
                    if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True):
                        if str_app.session_state[f"q_idx_{selected_sheet}"] < total_q - 1:
                            str_app.session_state[f"q_idx_{selected_sheet}"] += 1
                        else:
                            str_app.session_state[f"q_idx_{selected_sheet}"] = 0
                        scroll_to_top()
                        str_app.rerun()

        elif mode == "🔄 Ôn lại câu trả lời sai":
            str_app.title("🔄 Ôn Lại Câu Trả Lời Sai")
            str_app.markdown("---")
            
            wrong_list = str_app.session_state["wrong_questions"]
            if not wrong_list:
                str_app.info("🎉 Hiện tại bạn chưa có câu trả lời sai nào được lưu lại.")
            else:
                str_app.write(f"Bạn đang có **{len(wrong_list)}** câu cần ôn tập lại.")
                if "wrong_idx" not in str_app.session_state:
                    str_app.session_state["wrong_idx"] = 0
                    
                w_idx = str_app.session_state["wrong_idx"]
                if w_idx >= len(wrong_list):
                    w_idx = 0
                    str_app.session_state["wrong_idx"] = 0
                    
                w_item = wrong_list[w_idx]
                
                str_app.markdown(f"""
                    <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.05rem; font-weight: 700; color: #f43f5e;">⚠️ Nguồn: {w_item.get('sheet', 'N/A')}</span>
                            <span class="badge-topic">Câu sai {w_idx + 1} / {len(wrong_list)}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                str_app.markdown(f"""
                    <div class="question-card">
                        <div class="question-title">
                            <b>Câu {w_idx + 1}:</b> {w_item['question']}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                shuff_data = get_shuffled_options(w_item, f"shuff_wrong_{w_idx}")
                options = shuff_data["options"]
                correct_letter = shuff_data["correct"]
                
                w_storage_key = f"user_ans_wrong_{w_idx}"
                w_answered_key = f"answered_wrong_{w_idx}"
                is_w_answered = str_app.session_state.get(w_answered_key, False)
                
                w_default_idx = None
                if str_app.session_state.get(w_storage_key) in options:
                    w_default_idx = options.index(str_app.session_state.get(w_storage_key))
                    
                if not is_w_answered:
                    w_choice = str_app.radio("Lựa chọn đáp án:", options, index=w_default_idx, key=f"radio_wrong_{w_idx}", label_visibility="collapsed")
                    if w_choice is not None:
                        str_app.session_state[w_storage_key] = w_choice
                        str_app.session_state[w_answered_key] = True
                        
                        is_correct = w_choice.strip().upper().startswith(correct_letter)
                        update_spaced_repetition(w_item["question"], is_correct)
                        if is_correct:
                            wrong_list = [w for w in wrong_list if w.get("question") != w_item.get("question")]
                            str_app.session_state["wrong_questions"] = wrong_list
                            save_current_progress()
                        str_app.rerun()
                else:
                    saved_w_choice = str_app.session_state.get(w_storage_key)
                    str_app.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_w_choice}</b></p>", unsafe_allow_html=True)
                    
                    str_app.markdown("<br>", unsafe_allow_html=True)
                    if saved_w_choice.strip().upper().startswith(correct_letter):
                        str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}. (Đã xóa khỏi danh sách câu sai)")
                    else:
                        correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                        str_app.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
                
                str_app.markdown("<br>", unsafe_allow_html=True)
                if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_wrong_{w_idx}"):
                    if str_app.session_state["wrong_idx"] < len(wrong_list) - 1:
                        str_app.session_state["wrong_idx"] += 1
                    else:
                        str_app.session_state["wrong_idx"] = 0
                    scroll_to_top()
                    str_app.rerun()

        elif mode == "🧠 Spaced Repetition (Ôn thông minh)":
            str_app.title("🧠 Chế Độ Ôn Thông Minh (Spaced Repetition)")
            str_app.markdown("---")
            
            all_questions = []
            for sh, ql in sheets_data.items():
                all_questions.extend(ql)
                
            sr_dict = str_app.session_state["spaced_repetition_data"]
            current_time = time.time()
            
            due_questions = []
            for q in all_questions:
                q_txt = q["question"]
                if q_txt not in sr_dict:
                    due_questions.append(q)
                else:
                    if sr_dict[q_txt]["next_review"] <= current_time:
                        due_questions.append(q)
                        
            if not due_questions:
                str_app.success("🎉 Tuyệt vời! Hiện tại không có câu hỏi nào đến hạn ôn tập theo phương pháp Spaced Repetition. Hãy quay lại sau hoặc ôn tập theo chuyên đề nhé!")
            else:
                str_app.write(f"Đang có **{len(due_questions)}** câu đến lịch ôn tập thông minh cần giải quyết.")
                if "sr_idx" not in str_app.session_state:
                    str_app.session_state["sr_idx"] = 0
                    
                sr_idx = str_app.session_state["sr_idx"]
                if sr_idx >= len(due_questions):
                    sr_idx = 0
                    str_app.session_state["sr_idx"] = 0
                    
                sr_item = due_questions[sr_idx]
                box_level = sr_dict.get(sr_item["question"], {}).get("box", 1)
                
                str_app.markdown(f"""
                    <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">🧠 Hộp ghi nhớ (Level {box_level}/5) — Chuyên đề: {sr_item.get('sheet')}</span>
                            <span class="badge-topic">Câu thông minh {sr_idx + 1} / {len(due_questions)}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                str_app.markdown(f"""
                    <div class="question-card">
                        <div class="question-title">
                            <b>Câu hỏi:</b> {sr_item['question']}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                shuff_data = get_shuffled_options(sr_item, f"shuff_sr_{sr_idx}")
                options = shuff_data["options"]
                correct_letter = shuff_data["correct"]
                
                sr_storage_key = f"user_ans_sr_{sr_idx}"
                sr_answered_key = f"answered_sr_{sr_idx}"
                is_sr_answered = str_app.session_state.get(sr_answered_key, False)
                
                sr_default_idx = None
                if str_app.session_state.get(sr_storage_key) in options:
                    sr_default_idx = options.index(str_app.session_state.get(sr_storage_key))
                    
                if not is_sr_answered:
                    sr_choice = str_app.radio("Lựa chọn đáp án:", options, index=sr_default_idx, key=f"radio_sr_{sr_idx}", label_visibility="collapsed")
                    if sr_choice is not None:
                        str_app.session_state[sr_storage_key] = sr_choice
                        str_app.session_state[sr_answered_key] = True
                        
                        is_correct = sr_choice.strip().upper().startswith(correct_letter)
                        update_spaced_repetition(sr_item["question"], is_correct)
                        if not is_correct and not any(w.get("question") == sr_item["question"] for w in str_app.session_state["wrong_questions"]):
                            str_app.session_state["wrong_questions"].append(sr_item)
                        save_current_progress()
                        str_app.rerun()
                else:
                    saved_sr_choice = str_app.session_state.get(sr_storage_key)
                    str_app.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_sr_choice}</b></p>", unsafe_allow_html=True)
                    
                    str_app.markdown("<br>", unsafe_allow_html=True)
                    if saved_sr_choice.strip().upper().startswith(correct_letter):
                        str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}. (Đã nâng cấp cấp độ hộp ghi nhớ)")
                    else:
                        correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                        str_app.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**. (Đã đưa về hộp 1 để ôn lại)")
                
                str_app.markdown("<br>", unsafe_allow_html=True)
                if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_sr_{sr_idx}"):
                    if str_app.session_state["sr_idx"] < len(due_questions) - 1:
                        str_app.session_state["sr_idx"] += 1
                    else:
                        str_app.session_state["sr_idx"] = 0
                    scroll_to_top()
                    str_app.rerun()

        elif mode == "⭐ Tất cả câu hỏi cần ghi nhớ":
            str_app.title("⭐ Tất Cả Câu Hỏi Cần Ghi Nhớ")
            str_app.markdown("---")
            
            bm_list = str_app.session_state["bookmarked_questions"]
            if not bm_list:
                str_app.info("⭐ Bạn chưa đánh dấu câu hỏi nào cần ghi nhớ cả.")
            else:
                str_app.write(f"Tổng số câu bạn đã đánh dấu ghi nhớ: **{len(bm_list)}** câu.")
                if "global_bm_idx" not in str_app.session_state:
                    str_app.session_state["global_bm_idx"] = 0
                    
                gbm_idx = str_app.session_state["global_bm_idx"]
                if gbm_idx >= len(bm_list):
                    gbm_idx = 0
                    str_app.session_state["global_bm_idx"] = 0
                    
                bm_item = bm_list[gbm_idx]
                
                str_app.markdown(f"""
                    <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.05rem; font-weight: 700; color: #eab308;">⭐ Chuyên đề: {bm_item.get('sheet', 'N/A')}</span>
                            <span class="badge-topic">Đã lưu {gbm_idx + 1} / {len(bm_list)}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                col_bm1, col_bm2 = str_app.columns([4, 1])
                with col_bm2:
                    if str_app.button("❌ Bỏ lưu", key=f"remove_global_bm_{gbm_idx}__", use_container_width=True):
                        str_app.session_state["bookmarked_questions"] = [b for b in str_app.session_state["bookmarked_questions"] if b.get("question") != bm_item.get("question")]
                        save_current_progress()
                        str_app.toast("Đã xóa khỏi danh sách ghi nhớ!", icon="ℹ️")
                        str_app.rerun()

                str_app.markdown(f"""
                    <div class="question-card">
                        <div class="question-title">
                            <b>Câu hỏi:</b> {bm_item['question']}
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                options = bm_item["options"]
                if not options:
                    options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
                
                correct_letter = bm_item.get("correct", "A").strip().upper()
                
                str_app.markdown("##### 💡 Đáp án chuẩn:")
                for opt in options:
                    opt_letter = opt.strip().upper()[:1]
                    if opt_letter == correct_letter:
                        str_app.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
                    else:
                        str_app.markdown(f"- {opt}")
                
                str_app.markdown("<br>", unsafe_allow_html=True)
                if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_global_bm_{gbm_idx}"):
                    if str_app.session_state["global_bm_idx"] < len(bm_list) - 1:
                        str_app.session_state["global_bm_idx"] += 1
                    else:
                        str_app.session_state["global_bm_idx"] = 0
                    scroll_to_top()
                    str_app.rerun()

        elif mode == "📂 Ôn gộp tất cả (50 câu/phần)":
            str_app.title("📂 Ôn Gộp Tất Cả Chuyên Đề (50 Câu/Phần)")
            str_app.markdown("---")
            
            all_questions = []
            for sh, ql in sheets_data.items():
                all_questions.extend(ql)
                
            total_all = len(all_questions)
            if total_all == 0:
                str_app.warning("Không có dữ liệu câu hỏi.")
            else:
                chunk_size = 50
                total_chunks = (total_all // chunk_size) + (1 if total_all % chunk_size != 0 else 0)
                
                str_app.sidebar.markdown("""
                    <div class="metric-card-container" style="text-align: left;">
                        <div class="metric-label" style="margin-bottom: 8px;">📁 Chọn Phần Ôn Tập</div>
                """, unsafe_allow_html=True)
                
                chunk_names = []
                for i in range(total_chunks):
                    is_done = (i in str_app.session_state["completed_chunks"] or i in str_app.session_state["passed_tests"])
                    prefix = "✅ " if is_done else "📌 "
                    chunk_names.append(f"{prefix}Phần {i+1}")
                    
                selected_chunk_name = str_app.sidebar.selectbox("Chọn phần:", chunk_names, label_visibility="collapsed", key="sidebar_chunk_select")
                str_app.sidebar.markdown("</div>", unsafe_allow_html=True)
                
                c_num = int(selected_chunk_name.split("Phần")[1].strip()) - 1
                
                start_idx = c_num * chunk_size
                end_idx = min((c_num + 1) * chunk_size, total_all)
                
                if f"chunk_q_{c_num}" not in str_app.session_state:
                    chunk_qs = all_questions.copy()
                    random.seed(42)
                    random.shuffle(chunk_qs)
                    str_app.session_state[f"chunk_q_{c_num}"] = chunk_qs[start_idx:end_idx]
                    
                current_chunk_questions = str_app.session_state[f"chunk_q_{c_num}"]
                actual_chunk_len = len(current_chunk_questions)
                
                str_app.sidebar.markdown(f"""
                    <div class="metric-card-container">
                        <div class="metric-label">📊 Tổng số câu của phần</div>
                        <div class="metric-value">{actual_chunk_len}</div>
                    </div>
                """, unsafe_allow_html=True)
                
                if f"chunk_studied_count_{c_num}" not in str_app.session_state:
                    str_app.session_state[f"chunk_studied_count_{c_num}"] = set()

                studied_set = str_app.session_state[f"chunk_studied_count_{c_num}"]
                is_fully_studied = len(studied_set) >= actual_chunk_len

                sub_modes = ["📖 Ôn tập từng câu", "🔄 Làm lại phần này", "⚠️ Ôn các câu sai"]
                if is_fully_studied or (c_num in str_app.session_state["completed_chunks"]):
                    sub_modes.insert(1, "📝 Bài kiểm tra chốt kiến thức")
                
                sub_mode = str_app.radio("Chế độ học trong phần:", sub_modes, horizontal=True, label_visibility="collapsed")
                
                if sub_mode == "📝 Bài kiểm tra chốt kiến thức" and not is_fully_studied and (c_num not in str_app.session_state["completed_chunks"]):
                    str_app.warning("⚠️ Ông chưa ôn tập hết các câu trong phần này! Hãy hoàn thành phần 'Ôn tập từng câu' trước khi làm bài kiểm tra chốt.")
                    sub_mode = "📖 Ôn tập từng câu"

                str_app.markdown("---")
                
                if sub_mode == "📖 Ôn tập từng câu":
                    if f"gop_idx_{c_num}" not in str_app.session_state:
                        str_app.session_state[f"gop_idx_{c_num}"] = 0
                    
                    g_idx = str_app.session_state[f"gop_idx_{c_num}"]
                    
                    if g_idx >= actual_chunk_len:
                        str_app.success(f"🎉 Ông đã ôn tập xong toàn bộ {actual_chunk_len} câu của Phần {c_num + 1}! Bây giờ có thể chuyển sang Bài kiểm tra chốt kiến thức.")
                        str_app.markdown("---")
                        col_btn1, col_btn2 = str_app.columns(2)
                        with col_btn1:
                            if str_app.button("🔄 Ôn lại từ đầu", type="secondary", use_container_width=True):
                                str_app.session_state[f"gop_idx_{c_num}"] = 0
                                str_app.session_state[f"chunk_studied_count_{c_num}"] = set()
                                keys_to_del = [k for k in str_app.session_state.keys() if k.startswith(f"shuff_gop_{c_num}_") or k.startswith(f"user_ans_gop_{c_num}_") or k.startswith(f"answered_gop_{c_num}_")]
                                for k in keys_to_del:
                                    del str_app.session_state[k]
                                scroll_to_top()
                                str_app.rerun()
                        with col_btn2:
                            if str_app.button("📝 Mở khóa & Bắt đầu bài kiểm tra chốt", type="primary", use_container_width=True):
                                str_app.session_state[f"auto_switch_test_{c_num}"] = True
                                scroll_to_top()
                                str_app.rerun()
                    else:
                        if str_app.session_state.get(f"auto_switch_test_{c_num}", False):
                            str_app.session_state[f"auto_switch_test_{c_num}"] = False
                            str_app.rerun()

                    if g_idx < actual_chunk_len:
                        q_item = current_chunk_questions[g_idx]
                        
                        str_app.markdown(f"""
                            <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">📂 Phần {c_num + 1} — Chuyên đề: {q_item['sheet']}</span>
                                    <span class="badge-topic">Câu {g_idx + 1} / {actual_chunk_len} (Đã ôn: {len(studied_set)}/{actual_chunk_len})</span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        is_bm = any(b.get("question") == q_item["question"] for b in str_app.session_state["bookmarked_questions"])
                        bm_label = "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
                        
                        col_q_head1, col_q_head2 = str_app.columns([4, 1])
                        with col_q_head2:
                            if str_app.button(bm_label, key=f"bm_gop_{c_num}_{g_idx}", use_container_width=True):
                                if is_bm:
                                    str_app.session_state["bookmarked_questions"] = [b for b in str_app.session_state["bookmarked_questions"] if b.get("question") != q_item["question"]]
                                    str_app.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
                                else:
                                    str_app.session_state["bookmarked_questions"].append(q_item)
                                    str_app.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
                                save_current_progress()
                                str_app.rerun()
                                
                        str_app.markdown(f"""
                            <div class="question-card">
                                <div class="question-title">
                                    <b>Câu {g_idx + 1}:</b> {q_item['question']}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        shuff_data = get_shuffled_options(q_item, f"shuff_gop_{c_num}_{g_idx}")
                        options = shuff_data["options"]
                        correct_letter = shuff_data["correct"]
                        
                        gop_storage_key = f"user_ans_gop_{c_num}_{g_idx}"
                        gop_answered_key = f"answered_gop_{c_num}_{g_idx}"
                        is_gop_answered = str_app.session_state.get(gop_answered_key, False)
                        
                        gop_default_idx = None
                        if str_app.session_state.get(gop_storage_key) in options:
                            gop_default_idx = options.index(str_app.session_state.get(gop_storage_key))
                            
                        if not is_gop_answered:
                            g_choice = str_app.radio("Lựa chọn đáp án:", options, index=gop_default_idx, key=f"radio_gop_{c_num}_{g_idx}", label_visibility="collapsed")
                            
                            if g_choice is not None:
                                str_app.session_state[gop_storage_key] = g_choice
                                str_app.session_state[gop_answered_key] = True
                                str_app.session_state[f"chunk_studied_count_{c_num}"].add(g_idx)
                                
                                is_correct = g_choice.strip().upper().startswith(correct_letter)
                                update_spaced_repetition(q_item["question"], is_correct)
                                if not is_correct:
                                    if not any(w.get("question") == q_item["question"] for w in str_app.session_state["wrong_questions"]):
                                        str_app.session_state["wrong_questions"].append(q_item)
                                        save_current_progress()
                                str_app.rerun()
                        else:
                            saved_gop_choice = str_app.session_state.get(gop_storage_key)
                            str_app.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_gop_choice}</b></p>", unsafe_allow_html=True)
                            
                            is_correct = saved_gop_choice.strip().upper().startswith(correct_letter)
                            str_app.markdown("<br>", unsafe_allow_html=True)
                            if is_correct:
                                str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                            else:
                                correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                                str_app.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
                                
                        str_app.markdown("<br>", unsafe_allow_html=True)
                        col_prev, col_next = str_app.columns(2)
                        with col_prev:
                            if str_app.button("⬅️ Câu trước", use_container_width=True):
                                if str_app.session_state[f"gop_idx_{c_num}"] > 0:
                                    str_app.session_state[f"gop_idx_{c_num}"] -= 1
                                else:
                                    str_app.session_state[f"gop_idx_{c_num}"] = actual_chunk_len - 1
                                scroll_to_top()
                                str_app.rerun()
                        with col_next:
                            if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True):
                                str_app.session_state[f"chunk_studied_count_{c_num}"].add(g_idx)
                                if str_app.session_state[f"gop_idx_{c_num}"] < actual_chunk_len:
                                    str_app.session_state[f"gop_idx_{c_num}"] += 1
                                scroll_to_top()
                                str_app.rerun()

                elif sub_mode == "📝 Bài kiểm tra chốt kiến thức":
                    str_app.markdown(f"### Bài Kiểm Tra - Phần {c_num + 1} ({actual_chunk_len} câu)")
                    
                    submitted_key = f"submitted_test_{c_num}"
                    if submitted_key not in str_app.session_state:
                        str_app.session_state[submitted_key] = False

                    if f"test_start_time_{c_num}" not in str_app.session_state:
                        str_app.session_state[f"test_start_time_{c_num}"] = time.time()
                        
                    if not str_app.session_state[submitted_key]:
                        elapsed_sec = int(time.time() - str_app.session_state.get(f"test_start_time_{c_num}", time.time()))
                        rem_sec = max(0, 59 * 60 - elapsed_sec)

                        timer_code = f"""
                        <div style="
                            background: linear-gradient(135deg, rgba(244, 63, 94, 0.15), rgba(15, 23, 42, 0.8));
                            border: 1px solid #f43f5e;
                            border-radius: 12px;
                            padding: 12px 20px;
                            text-align: center;
                            font-family: sans-serif;
                            box-shadow: 0 0 15px rgba(244, 63, 94, 0.2);
                            margin-bottom: 25px;
                        ">
                            <span style="color: #f8fafc; font-size: 1.1rem; font-weight: 600;">⏱️ Thời gian kiểm tra còn lại: </span>
                            <span id="countdown_chunk_{c_num}" style="color: #fb7185; font-size: 1.6rem; font-weight: 800; font-family: monospace;">--:--</span>
                        </div>
                        <script>
                            var duration = {rem_sec};
                            var display = document.querySelector('#countdown_chunk_{c_num}');
                            
                            function updateTimer() {{
                                var minutes = parseInt(duration / 60, 10);
                                var seconds = parseInt(duration % 60, 10);

                                minutes = minutes < 10 ? "0" + minutes : minutes;
                                seconds = seconds < 10 ? "0" + seconds : seconds;

                                display.textContent = minutes + ":" + seconds;

                                if (duration <= 0) {{
                                    display.textContent = "00:00 - HẾT GIỜ!";
                                    display.style.color = "#ff4d4d";
                                }} else {{
                                    duration--;
                                }}
                            }}
                            
                            updateTimer();
                            setInterval(updateTimer, 1000);
                        </script>
                        """
                        components.html(timer_code, height=75)

                        user_answers = {}
                        for i, q in enumerate(current_chunk_questions):
                            str_app.markdown(f"""
                                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                                    <span class="badge-topic">{q['sheet']}</span>
                                    <span style="color: #94a3b8; font-weight: 600;">Câu {i+1}</span>
                                </div>
                                <div class="question-card">
                                    <div class="question-title">
                                        {q['question']}
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                            
                            shuff_data = get_shuffled_options(q, f"shuff_test_{c_num}_{i}")
                            options = shuff_data["options"]
                            
                            test_ans_key = f"test_chunk_{c_num}_{i}"
                            ans = str_app.radio("Lựa chọn đáp án:", options, index=(options.index(str_app.session_state[test_ans_key]) if str_app.session_state.get(test_ans_key) in options else None), key=test_ans_key, label_visibility="collapsed")
                            user_answers[i] = ans
                            str_app.markdown("<hr style='margin: 30px 0; border-color: rgba(56, 189, 248, 0.2);'>", unsafe_allow_html=True)
                            
                        if str_app.button("📤 Nộp bài kiểm tra ngay", type="primary", use_container_width=True):
                            unanswered = [i + 1 for i in range(actual_chunk_len) if user_answers[i] is None]
                            if unanswered:
                                str_app.error(f"⚠️ Ông chưa làm xong tất cả các câu! Còn thiếu các câu: {', '.join(map(str, unanswered))}.")
                            else:
                                correct_count = 0
                                wrong_count = 0
                                for i, q in enumerate(current_chunk_questions):
                                    selected = user_answers[i]
                                    shuff_data = str_app.session_state.get(f"shuff_test_{c_num}_{i}", {"correct": "A"})
                                    correct_letter = shuff_data["correct"]
                                    
                                    is_correct = selected and selected.strip().upper().startswith(correct_letter)
                                    update_spaced_repetition(q["question"], is_correct)
                                    if is_correct:
                                        correct_count += 1
                                    else:
                                        wrong_count += 1
                                        if not any(w.get("question") == q["question"] for w in str_app.session_state["wrong_questions"]):
                                            str_app.session_state["wrong_questions"].append(q)
                                            
                                str_app.session_state[f"result_correct_{c_num}"] = correct_count
                                str_app.session_state[f"result_wrong_{c_num}"] = wrong_count
                                str_app.session_state[f"test_user_answers_{c_num}"] = user_answers
                                str_app.session_state[submitted_key] = True
                                str_app.session_state["completed_chunks"].add(c_num)
                                str_app.session_state["passed_tests"].add(c_num)
                                save_current_progress()
                                str_app.balloons()
                                scroll_to_top()
                                str_app.rerun()
                    else:
                        c_correct = str_app.session_state.get(f"result_correct_{c_num}", 0)
                        c_wrong = str_app.session_state.get(f"result_wrong_{c_num}", 0)
                        score_percent = (c_correct / actual_chunk_len) * 100
                        
                        str_app.success("🎉 Đã nộp bài kiểm tra thành công!")
                        str_app.markdown("### 📊 Kết Quả Bài Kiểm Tra")
                        col1, col2, col3 = str_app.columns(3)
                        col1.metric("Số câu đúng", f"{c_correct}/{actual_chunk_len}", f"{score_percent:.1f}%")
                        col2.metric("Số câu sai", f"{c_wrong}/{actual_chunk_len}")
                        col3.metric("Trạng thái", "Đã hoàn thành ✅")
                        
                        str_app.markdown("---")
                        str_app.markdown("### 🔍 Xem Lại Chi Tiết Các Câu Trả Lời Sai")
                        user_answers = str_app.session_state.get(f"test_user_answers_{c_num}", {})
                        wrong_items_in_test = []
                        for i, q in enumerate(current_chunk_questions):
                            selected = user_answers.get(i)
                            shuff_data = str_app.session_state.get(f"shuff_test_{c_num}_{i}", {"correct": "A"})
                            correct_letter = shuff_data["correct"]
                            
                            is_correct = selected and selected.strip().upper().startswith(correct_letter)
                            if not is_correct:
                                wrong_items_in_test.append((i, q, selected, shuff_data))
                                
                        if not wrong_items_in_test:
                            str_app.info("🎉 Tuyệt vời! Ông đã trả lời đúng tất cả các câu trong phần này.")
                        else:
                            for q_idx, q_item, user_sel, shuff_info in wrong_items_in_test:
                                str_app.markdown(f"**Câu {q_idx + 1}** *(Thuộc chuyên đề: {q_item['sheet']})*")
                                str_app.markdown(f"""
                                    <div class="question-card">
                                        <div class="question-title">{q_item['question']}</div>
                                    </div>
                                """, unsafe_allow_html=True)
                                correct_letter = shuff_info["correct"]
                                for opt in shuff_info["options"]:
                                    opt_letter = opt.strip().upper()[:1]
                                    is_this_correct = (opt_letter == correct_letter)
                                    is_user_chosen = (user_sel and opt.strip() == user_sel.strip())
                                    
                                    if is_this_correct:
                                        str_app.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
                                    elif is_user_chosen:
                                        str_app.markdown(f"- ❌ ~~{opt}~~ *(Ông đã chọn)*")
                                    else:
                                        str_app.markdown(f"- {opt}")
                                str_app.markdown("<hr style='margin: 20px 0; border-color: rgba(56, 189, 248, 0.2);'>", unsafe_allow_html=True)
                                
                        if str_app.button("🔄 Làm lại bài kiểm tra này", type="primary", use_container_width=True):
                            str_app.session_state[submitted_key] = False
                            str_app.session_state[f"test_start_time_{c_num}"] = time.time()
                            keys_to_del = [k for k in str_app.session_state.keys() if k.startswith(f"shuff_test_{c_num}_") or k.startswith(f"test_chunk_{c_num}_")]
                            for k in keys_to_del:
                                del str_app.session_state[k]
                            if f"test_user_answers_{c_num}" in str_app.session_state:
                                del str_app.session_state[f"test_user_answers_{c_num}"]
                            if c_num in str_app.session_state["completed_chunks"]:
                                str_app.session_state["completed_chunks"].remove(c_num)
                            if c_num in str_app.session_state["passed_tests"]:
                                str_app.session_state["passed_tests"].remove(c_num)
                            save_current_progress()
                            scroll_to_top()
                            str_app.rerun()

                elif sub_mode == "🔄 Làm lại phần này":
                    str_app.session_state[f"gop_idx_{c_num}"] = 0
                    str_app.session_state[f"chunk_studied_count_{c_num}"] = set()
                    keys_to_del = [k for k in str_app.session_state.keys() if k.startswith(f"shuff_gop_{c_num}_") or k.startswith(f"shuff_test_{c_num}_") or k.startswith(f"user_ans_gop_{c_num}_") or k.startswith(f"answered_gop_{c_num}_") or k.startswith(f"test_chunk_{c_num}_")]
                    for k in keys_to_del:
                        del str_app.session_state[k]
                    if f"submitted_test_{c_num}" in str_app.session_state:
                        str_app.session_state[f"submitted_test_{c_num}"] = False
                    if f"test_user_answers_{c_num}" in str_app.session_state:
                        del str_app.session_state[f"test_user_answers_{c_num}"]
                    if f"test_start_time_{c_num}" in str_app.session_state:
                        del str_app.session_state[f"test_start_time_{c_num}"]
                    if c_num in str_app.session_state["completed_chunks"]:
                        str_app.session_state["completed_chunks"].remove(c_num)
                    if c_num in str_app.session_state["passed_tests"]:
                        str_app.session_state["passed_tests"].remove(c_num)
                    save_current_progress()
                    str_app.success(f"Đã reset Phần {c_num + 1} thành công!")
                    scroll_to_top()
                    str_app.rerun()

                elif sub_mode == "⚠️ Ôn các câu sai":
                    str_app.markdown(f"### Ôn Lại Các Câu Sai Trong Phần {c_num + 1}")
                    chunk_questions_texts = set(q["question"] for q in current_chunk_questions)
                    wrong_in_chunk = [q for q in str_app.session_state["wrong_questions"] if q.get("question") in chunk_questions_texts]
                    
                    if not wrong_in_chunk:
                        str_app.info("🎉 Phần này bạn chưa trả lời sai câu nào cả!")
                    else:
                        if f"wrong_chunk_idx_{c_num}" not in str_app.session_state:
                            str_app.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            
                        wc_idx = str_app.session_state[f"wrong_chunk_idx_{c_num}"]
                        if wc_idx >= len(wrong_in_chunk):
                            wc_idx = 0
                            str_app.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            
                        wc_item = wrong_in_chunk[wc_idx]
                        
                        str_app.markdown(f"""
                            <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 1.05rem; font-weight: 700; color: #f43f5e;">⚠️ Ôn câu sai Phần {c_num + 1}</span>
                                    <span class="badge-topic">Câu sai {wc_idx + 1} / {len(wrong_in_chunk)}</span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        str_app.markdown(f"""
                            <div class="question-card">
                                <div class="question-title">
                                    <b>Câu hỏi:</b> {wc_item['question']}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        shuff_data = get_shuffled_options(wc_item, f"shuff_wc_{c_num}_{wc_idx}")
                        options = shuff_data["options"]
                        correct_letter = shuff_data["correct"]
                        
                        wc_storage_key = f"user_ans_wc_{c_num}_{wc_idx}"
                        wc_answered_key = f"answered_wc_{c_num}_{wc_idx}"
                        is_wc_answered = str_app.session_state.get(wc_answered_key, False)
                        
                        wc_default_idx = None
                        if str_app.session_state.get(wc_storage_key) in options:
                            wc_default_idx = options.index(str_app.session_state.get(wc_storage_key))
                            
                        if not is_wc_answered:
                            wc_choice = str_app.radio("Lựa chọn đáp án:", options, index=wc_default_idx, key=f"radio_wc_{c_num}_{wc_idx}", label_visibility="collapsed")
                            if wc_choice is not None:
                                str_app.session_state[wc_storage_key] = wc_choice
                                str_app.session_state[wc_answered_key] = True
                                
                                is_correct = wc_choice.strip().upper().startswith(correct_letter)
                                update_spaced_repetition(wc_item["question"], is_correct)
                                if is_correct:
                                    str_app.session_state["wrong_questions"] = [
                                        w for w in str_app.session_state["wrong_questions"] 
                                        if w.get("question") != wc_item.get("question")
                                    ]
                                    save_current_progress()
                                str_app.rerun()
                        else:
                            saved_wc_choice = str_app.session_state.get(wc_storage_key)
                            str_app.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_wc_choice}</b></p>", unsafe_allow_html=True)
                            
                            str_app.markdown("<br>", unsafe_allow_html=True)
                            if saved_wc_choice.strip().upper().startswith(correct_letter):
                                str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}. (Đã xóa khỏi danh sách câu sai)")
                            else:
                                correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                                str_app.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
                        
                        str_app.markdown("<br>", unsafe_allow_html=True)
                        if str_app.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_wc_{c_num}_{wc_idx}"):
                            if str_app.session_state[f"wrong_chunk_idx_{c_num}"] < len(wrong_in_chunk) - 1:
                                str_app.session_state[f"wrong_chunk_idx_{c_num}"] += 1
                            else:
                                str_app.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            scroll_to_top()
                            str_app.rerun()

        elif mode == "📝 Thi thử (Mock Test)":
            str_app.title("📝 Chế Độ Thi Thử (Mock Test)")
            str_app.markdown("---")
            
            all_questions = []
            for sh, ql in sheets_data.items():
                all_questions.extend(ql)
                
            if "mock_started" not in str_app.session_state:
                str_app.session_state["mock_started"] = False
                
            if not str_app.session_state["mock_started"]:
                str_app.markdown("""
                    <div class="welcome-card" style="text-align: left; padding: 30px;">
                        <h3>📋 Thông tin bài thi thử mô phỏng:</h3>
                        <p style="color: #cbd5e1; font-size: 1.1rem; margin-top: 10px;">- Bài thi gồm <b>50 câu hỏi ngẫu nhiên</b> được trộn đều từ toàn bộ ngân hàng câu hỏi an toàn điện.</p>
                        <p style="color: #cbd5e1; font-size: 1.1rem;">- Thời gian làm bài chính thức: <b>59 phút</b> (có đếm ngược trực tiếp).</p>
                        <p style="color: #cbd5e1; font-size: 1.1rem;">- Đánh giá chính xác năng lực và mức độ sẵn sàng trước kỳ thi chính thức.</p>
                    </div>
                """, unsafe_allow_html=True)
                str_app.markdown("<br>", unsafe_allow_html=True)
                
                col_m1, col_m2, col_m3 = str_app.columns([1, 1.5, 1])
                with col_m2:
                    if str_app.button("🚀 Bắt đầu làm bài thi ngay", type="primary", use_container_width=True):
                        str_app.session_state["mock_started"] = True
                        str_app.session_state["mock_start_time"] = time.time()
                        str_app.session_state["mock_questions"] = random.sample(all_questions, min(50, len(all_questions)))
                        str_app.session_state["mock_answers"] = {}
                        keys_to_del = [k for k in str_app.session_state.keys() if k.startswith("shuff_mock_") or k.startswith("mock_q_")]
                        for k in keys_to_del:
                            del str_app.session_state[k]
                        scroll_to_top()
                        str_app.rerun()
            else:
                elapsed_sec = int(time.time() - str_app.session_state.get("mock_start_time", time.time()))
                rem_sec = max(0, 59 * 60 - elapsed_sec)

                timer_code = f"""
                <div style="
                    background: linear-gradient(135deg, rgba(244, 63, 94, 0.15), rgba(15, 23, 42, 0.8));
                    border: 1px solid #f43f5e;
                    border-radius: 12px;
                    padding: 12px 20px;
                    text-align: center;
                    font-family: sans-serif;
                    box-shadow: 0 0 15px rgba(244, 63, 94, 0.2);
                ">
                    <span style="color: #f8fafc; font-size: 1.1rem; font-weight: 600;">⏱️ Thời gian còn lại: </span>
                    <span id="countdown" style="color: #fb7185; font-size: 1.6rem; font-weight: 800; font-family: monospace;">--:--</span>
                </div>
                <script>
                    var duration = {rem_sec};
                    var display = document.querySelector('#countdown');
                    
                    function updateTimer() {{
                        var minutes = parseInt(duration / 60, 10);
                        var seconds = parseInt(duration % 60, 10);

                        minutes = minutes < 10 ? "0" + minutes : minutes;
                        seconds = seconds < 10 ? "0" + seconds : seconds;

                        display.textContent = minutes + ":" + seconds;

                        if (duration <= 0) {{
                            display.textContent = "00:00 - HẾT GIỜ!";
                            display.style.color = "#ff4d4d";
                        }} else {{
                            duration--;
                        }}
                    }}
                    
                    updateTimer();
                    setInterval(updateTimer, 1000);
                </script>
                """
                components.html(timer_code, height=75)

                mock_qs = str_app.session_state["mock_questions"]
                str_app.write(f"Đang làm bài thi thử chính thức ({len(mock_qs)} câu).")
                str_app.markdown("---")
                
                for i, q in enumerate(mock_qs):
                    str_app.markdown(f"""
                        <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                            <span class="badge-topic">{q['sheet']}</span>
                            <span style="color: #94a3b8; font-weight: 600;">Câu {i+1} / {len(mock_qs)}</span>
                        </div>
                        <div class="question-card">
                            <div class="question-title">
                                {q['question']}
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    shuff_data = get_shuffled_options(q, f"shuff_mock_{i}")
                    options = shuff_data["options"]
                    
                    mock_ans_key = f"mock_q_{i}"
                    ans = str_app.radio("Lựa chọn đáp án:", options, index=(options.index(str_app.session_state[mock_ans_key]) if str_app.session_state.get(mock_ans_key) in options else None), key=mock_ans_key, label_visibility="collapsed")
                    str_app.session_state["mock_answers"][i] = ans
                    str_app.markdown("<hr style='margin: 30px 0; border-color: rgba(56, 189, 248, 0.2);'>", unsafe_allow_html=True)
                    
                if str_app.button("📤 Nộp bài thi thử", type="primary", use_container_width=True):
                    for i, q in enumerate(mock_qs):
                        selected = str_app.session_state["mock_answers"].get(i)
                        shuff_data = str_app.session_state.get(f"shuff_mock_{i}", {"correct": "A"})
                        correct_letter = shuff_data["correct"]
                        is_correct = selected and selected.strip().upper().startswith(correct_letter)
                        update_spaced_repetition(q["question"], is_correct)
                    str_app.success("Đã nộp bài và cập nhật hệ thống Spaced Repetition thành công!")
                    str_app.balloons()
                    scroll_to_top()
                    if str_app.button("Làm bài thi mới", use_container_width=True):
                        str_app.session_state["mock_started"] = False
                        scroll_to_top()
                        str_app.rerun()

        elif mode == "⚙️ Quản lý kho lưu trữ & Dữ liệu":
            str_app.title("⚙️ Trung Tâm Quản Lý Trạng Thái & Kho Lưu Trữ")
            str_app.markdown("---")
            
            str_app.markdown("""
                <div class="welcome-card" style="text-align: left; padding: 25px;">
                    <h3>💾 Quản lý bộ nhớ ứng dụng & Tiến độ học tập</h3>
                    <p style="color: #cbd5e1; font-size: 1.05rem; margin-top: 10px;">
                        Hệ thống tự động lưu trữ toàn bộ tiến độ, lịch sử ôn tập, danh sách câu sai và trạng thái Spaced Repetition vào tệp cục bộ (<code>quiz_progress.json</code>).
                        Bạn có thể tải xuống tệp dữ liệu để dự phòng hoặc tải lên để khôi phục bất cứ lúc nào.
                    </p>
                </div>
            """, unsafe_allow_html=True)
            str_app.markdown("<br>", unsafe_allow_html=True)
            
            col_st1, col_st2 = str_app.columns(2)
            
            with col_st1:
                str_app.markdown("### 📥 Sao lưu & Xuất dữ liệu")
                str_app.markdown("Tải tệp tiến độ hiện tại về máy tính để làm bản sao lưu an toàn.")
                
                # Tạo chuỗi JSON từ trạng thái hiện tại
                current_saved_data = load_saved_progress()
                json_str = json.dumps(current_saved_data, ensure_ascii=False, indent=4)
                
                str_app.download_button(
                    label="📥 Tải xuống tệp tiến độ (.json)",
                    data=json_str,
                    file_name="quiz_progress_backup.json",
                    mime="application/json",
                    type="primary",
                    use_container_width=True
                )
                
            with col_st2:
                str_app.markdown("### 📤 Khôi phục & Nhập dữ liệu")
                str_app.markdown("Tải lên tệp sao lưu `.json` để khôi phục toàn bộ tiến độ học trước đó.")
                
                uploaded_file = str_app.file_uploader("Chọn tệp sao lưu JSON:", type=["json"], label_visibility="collapsed")
                if uploaded_file is not None:
                    try:
                        imported_data = json.load(uploaded_file)
                        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                            json.dump(imported_data, f, ensure_ascii=False, indent=4)
                        str_app.success("✅ Khôi phục dữ liệu thành công! Hãy tải lại trang để áp dụng.")
                        if str_app.button("🔄 Tải lại ứng dụng ngay", use_container_width=True):
                            str_app.rerun()
                    except Exception as e:
                        str_app.error(f"❌ Tệp JSON không hợp lệ: {str(e)}")
            
            str_app.markdown("---")
            str_app.markdown("### ⚠️ Vùng Nguy Hiểm (Quản lý trạng thái nhanh)")
            
            col_d1, col_d2, col_d3 = str_app.columns(3)
            
            with col_d1:
                if str_app.button("🧹 Xóa danh sách câu sai", use_container_width=True):
                    str_app.session_state["wrong_questions"] = []
                    save_current_progress()
                    str_app.toast("Đã dọn sạch danh sách câu sai!", icon="🗑️")
                    str_app.rerun()
                    
            with col_d2:
                if str_app.button("⭐ Xóa tất cả câu đã lưu", use_container_width=True):
                    str_app.session_state["bookmarked_questions"] = []
                    save_current_progress()
                    str_app.toast("Đã xóa danh sách ghi nhớ!", icon="🗑️")
                    str_app.rerun()
                    
            with col_d3:
                if str_app.button("🔄 Reset toàn bộ tiến độ", type="primary", use_container_width=True):
                    if os.path.exists(PROGRESS_FILE):
                        os.remove(PROGRESS_FILE)
                    for key in list(str_app.session_state.keys()):
                        del str_app.session_state[key]
                    str_app.success("Đã thiết lập lại toàn bộ ứng dụng từ đầu!")
                    time.sleep(1)
                    str_app.rerun()
                    # --- THÊM PHẦN KẾT THÚC ÔN TẬP VÀ GỬI MAIL VỀ GMAIL ---

      # 2. Tstr_app.markdown("---")
# Chỉ hiển thị nút này khi người dùng thực sự đang trong phiên ôn tập
if str_app.session_state.get("started", False):
  str_app.markdown("---")
  str_app.markdown("### 🏁 Hoàn thành phiên ôn tập")

  if str_app.button(
      "🏠 Hoàn thành phiên ôn tập & Về màn hình Welcome",
      type="primary",
      use_container_width=True,
      key="unique_btn_return_welcome_final_2026",
  ):
    with str_app.spinner(
        "Đang gửi báo cáo tổng kết về Gmail và về trang chủ..."
    ):
      summary_text = (
          f"Chào Đức Anh,\n\n"
          f"Hệ thống ghi nhận ông vừa hoàn thành một phiên ôn tập chuẩn bị thi Điện lực Điện Biên:\n"
          f"- Thời gian kết thúc phiên: Hôm nay\n"
          f"- Trạng thái: Đã hoàn thành phiên ôn luyện và reset phiên làm việc để về trang chủ.\n\n"
          f"Chúc ông tiếp tục giữ vững phong độ cho kỳ thi sắp tới!"
      )

      try:
        # send_daily_reminder_email()
        str_app.success(
            "Đã gửi báo cáo tự động về Gmail thành công! Đang chuyển trang..."
        )
      except Exception as e:
        str_app.error(f"Lỗi gửi email tự động: {e}")

      for key in list(str_app.session_state.keys()):
        del str_app.session_state[key]

      time.sleep(1)
      str_app.rerun()
       # Đưa nút hoàn thành lên thanh sidebar (thanh task bên trái)
# Chỉ hiển thị nút chức năng trên sidebar khi người dùng đã vào phiên ôn tập
    if str_app.session_state.get("is_studying", False):
      if str_app.sidebar.button(
          "🏠 Hoàn thành & Về Welcome",
          type="primary",
          use_container_width=True,
          key="btn_sidebar_return_welcome",
      ):
        with str_app.sidebar.spinner("Đang gửi báo cáo và về trang chủ..."):
          # 1. Soạn nội dung báo cáo tự động
          summary_text = (
              f"Chào Đức Anh,\n\n"
              f"Hệ thống ghi nhận ông vừa hoàn thành phiên ôn tập Điện lực Điện Biên từ Sidebar:\n"
              f"- Trạng thái: Đã reset phiên làm việc về trang chủ.\n"
              f"Chúc ông đạt kết quả cao trong kỳ thi sắp tới!"
          )
          
          # Phần gửi mail và reset session giữ nguyên ở đây...

    try:
      # 2. Sử dụng str_app.session_state thay vì st.session_state để đồng bộ với app của ông
      send_daily_reminder_email(
          receiver_email="ducanhdao74@gmail.com",
          completed_questions_count=str_app.session_state.get(
              "completed_count", 0
          ),
          total_questions_count=str_app.session_state.get(
              "total_questions", 1000
          ),
          bookmarked_count=str_app.session_state.get("bookmarked_count", 0),
          total_study_hours=str_app.session_state.get("study_hours", 1.0),
          is_completion=True,  # <--- Thêm đúng dòng này vào đây
      )
      str_app.sidebar.success("Đã gửi báo cáo tự động về Gmail thành công!")
    except Exception as e:
      str_app.sidebar.error(f"Lỗi gửi email: {e}")

    # 3. Dừng 2 giây để chắc chắn mail đã bắn đi thành công
    time.sleep(2)

    # 4. Xóa sạch session state và quay về màn hình Welcome
    for key in list(str_app.session_state.keys()):
      del str_app.session_state[key]

    str_app.rerun()
