from datetime import datetime
import json
import os
import random
import time
import openpyxl
import pandas as pd
import streamlit as str_app
import streamlit.components.v1 as components
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

str_app.set_page_config(
    page_title="⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện",
    page_icon="⚡",
    layout="wide",
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
      "completed_chunks": [],
      "passed_tests": [],
      "bookmarked_questions": [],
      "wrong_questions": [],
      "spaced_repetition_data": {},
      "total_study_seconds": 0,
      "last_login_date": "",
  }


def save_current_progress():
  if "start_session_time" in str_app.session_state:
    elapsed = time.time() - str_app.session_state["start_session_time"]
    str_app.session_state["start_session_time"] = time.time()
    saved_data = load_saved_progress()
    saved_data["total_study_seconds"] = (
        saved_data.get("total_study_seconds", 0) + elapsed
    )
  else:
    saved_data = load_saved_progress()

  data = {
      "completed_chunks": list(
          str_app.session_state.get("completed_chunks", [])
      ),
      "passed_tests": list(str_app.session_state.get("passed_tests", [])),
      "bookmarked_questions": str_app.session_state.get(
          "bookmarked_questions", []
      ),
      "wrong_questions": str_app.session_state.get("wrong_questions", []),
      "spaced_repetition_data": str_app.session_state.get(
          "spaced_repetition_data", {}
      ),
      "total_study_seconds": saved_data.get("total_study_seconds", 0),
      "last_login_date": str_app.session_state.get("last_login_date", ""),
  }
  try:
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=4)
    return True
  except:
    return False


def scroll_to_top():
  components.html(
      """
        <script>
            const doc = window.parent.document;
            setTimeout(() => {
                const main = doc.querySelector('.main') || doc.documentElement;
                main.scrollTo({top: 0, behavior: 'auto'});
                window.parent.scrollTo({top: 0, behavior: 'auto'});
            }, 10);
        </script>
    """,
      height=0,
  )


def send_daily_reminder_email(
    receiver_email,
    completed_questions_count,
    total_questions_count,
    bookmarked_count,
    total_study_hours,
):
  subject = "⚡ Nhắc nhở ôn tập An Toàn Điện mỗi ngày!"
  if total_study_hours < 1:
    time_str = f"{int(total_study_hours * 60)} phút"
  else:
    time_str = f"{total_study_hours:.1f} giờ"

  body = f"""
Chào Đức Anh,

Hôm nay là một ngày mới rồi! Hãy dành ra chút thời gian để vào ôn tập ngân hàng câu hỏi An Toàn Điện nhé:

📊 Tiến độ hiện tại của ông:
- Số câu đã hoàn thành: {completed_questions_count}/{total_questions_count} câu
- Tổng thời gian đã ôn tập: {time_str}
- Số câu hỏi đang cần ghi nhớ (Star): {bookmarked_count} câu

Chúc ông ôn thi thật tốt và đạt kết quả cao!
"""
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
    return None, "Không tìm thấy file Excel trong thư mục!"
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
              if any(
                  x in rgb_str for x in ["FF0000", "ED1C24", "C00000", "RED"]
              ):
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
            if hasattr(c, "indexed") and c.indexed in [10, 2]:
              is_red = True

          if val_str.lower().startswith("câu"):
            if current_q:
              questions.append({
                  "question": current_q,
                  "options": current_opts,
                  "correct": correct_ans,
                  "sheet": sheet,
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
            "sheet": sheet,
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
  str_app.session_state["wrong_questions"] = saved_prog.get(
      "wrong_questions", []
  )
if "bookmarked_questions" not in str_app.session_state:
  str_app.session_state["bookmarked_questions"] = saved_prog.get(
      "bookmarked_questions", []
  )
if "spaced_repetition_data" not in str_app.session_state:
  str_app.session_state["spaced_repetition_data"] = saved_prog.get(
      "spaced_repetition_data", {}
  )
if "completed_chunks" not in str_app.session_state:
  str_app.session_state["completed_chunks"] = set(
      saved_prog.get("completed_chunks", [])
  )
if "passed_tests" not in str_app.session_state:
  str_app.session_state["passed_tests"] = set(saved_prog.get("passed_tests", []))
if "app_started" not in str_app.session_state:
  str_app.session_state["app_started"] = False

save_current_progress()


def update_spaced_repetition(q_text, is_correct):
  sr_data = str_app.session_state["spaced_repetition_data"]
  now = time.time()
  if q_text not in sr_data:
    sr_data[q_text] = {"box": 1, "next_review": 0}

  item = sr_data[q_text]
  if is_correct:
    item["box"] = min(5, item["box"] + 1)
    intervals = [0, 0, 4 * 3600, 24 * 3600, 3 * 24 * 3600, 7 * 24 * 3600]
    item["next_review"] = now + intervals[item["box"]]
  else:
    item["box"] = 1
    item["next_review"] = now
  save_current_progress()


def scheduled_job():
  saved_data = load_saved_progress()
  bm_cnt = len(saved_data.get("bookmarked_questions", []))
  completed_chunks_set = set(saved_data.get("completed_chunks", [])).union(
      set(saved_data.get("passed_tests", []))
  )
  comp_q_cnt = min(len(completed_chunks_set) * 50, total_all_questions)
  total_hours = saved_data.get("total_study_seconds", 0) / 3600.0
  send_daily_reminder_email(
      SENDER_EMAIL,
      comp_q_cnt,
      total_all_questions,
      bm_cnt,
      total_hours,
  )


@str_app.cache_resource
def start_scheduler():
  scheduler = BackgroundScheduler()
  scheduler.add_job(scheduled_job, "cron", hour=7, minute=0)
  scheduler.start()
  return scheduler


try:
  start_scheduler()
except Exception:
  pass

# --- CYBERPUNK / DARK TECH GLASSMORPHISM CSS & OPTION CARDS ---
str_app.markdown(
    """
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
""",
    unsafe_allow_html=True,
)


def get_shuffled_options(q_item, session_key):
  if session_key not in str_app.session_state:
    orig_options = q_item["options"]
    if not orig_options:
      orig_options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

    correct_letter = q_item.get("correct", "A").strip().upper()

    opts_with_status = []
    for opt in orig_options:
      opt_prefix = opt[:1].strip().upper()
      is_corr = opt_prefix == correct_letter
      clean_text = (
          opt[2:].strip() if len(opt) > 2 and opt[1] in [".", ")"] else opt
      )
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
        "correct": new_correct_letter,
    }

  return str_app.session_state[session_key]


if not str_app.session_state["app_started"]:
  str_app.markdown("<br><br>", unsafe_allow_html=True)
  col_w1, col_w2, col_w3 = str_app.columns([1, 2.2, 1])
  with col_w2:
    str_app.markdown(
        """
            <div class="welcome-card">
                <h1 class="sparkle-title">⚡ Chào Đức Anh!</h1>
                <p style="font-size: 1.25rem; color: #e2e8f0; margin-top: 15px; font-weight: 500;">Hệ thống Ngân hàng câu hỏi An Toàn Điện đã sẵn sàng.</p>
                <p style="font-size: 1.05rem; color: #94a3b8; margin-top: 8px;">Chúc ông ôn tập thật tập trung, nắm trọn kiến thức và đạt kết quả cao nhất! 🚀</p>
            </div>
        """,
        unsafe_allow_html=True,
    )
    str_app.markdown("<br>", unsafe_allow_html=True)

    col_btn_center1, col_btn_center2, col_btn_center3 = str_app.columns(
        [1, 1.5, 1]
    )
    with col_btn_center2:
      if str_app.button(
          "✨ Bắt đầu vào ôn tập ngay", type="primary", use_container_width=True
      ):
        str_app.session_state["app_started"] = True
        scroll_to_top()
        str_app.rerun()
else:
  str_app.sidebar.markdown(
      """
        <div style="font-size: 1.2rem; font-weight: 800; color: #f8fafc; margin-bottom: 15px; display: flex; align-items: center; gap: 8px;">
            ⚡ Dashboard Tổng Quan
        </div>
    """,
      unsafe_allow_html=True,
  )

  completed_chunks_set = str_app.session_state["completed_chunks"].union(
      str_app.session_state["passed_tests"]
  )
  chunk_size_calc = 50
  completed_est_count = min(
      len(completed_chunks_set) * chunk_size_calc, total_all_questions
  )
  progress_ratio = (
      completed_est_count / total_all_questions
      if total_all_questions > 0
      else 0.0
  )

  current_total_seconds = saved_prog.get("total_study_seconds", 0) + (
      time.time() - str_app.session_state["start_session_time"]
  )
  current_total_hours = current_total_seconds / 3600.0
  bookmarked_count = len(str_app.session_state.get("bookmarked_questions", []))
  wrong_count = len(str_app.session_state.get("wrong_questions", []))

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
                <div class="metric-label" style="font-size: 0.7rem;">⚠️ Câu sai</div>
                <div class="metric-value" style="font-size: 1.2rem; color: #f43f5e; text-shadow: 0 0 10px rgba(244, 63, 94, 0.4);">{wrong_count}</div>
            </div>
        """,
        unsafe_allow_html=True,
    )

  str_app.sidebar.markdown(
      f"""
        <div class="metric-card-container">
            <div class="metric-label">⏱️ Tổng thời gian ôn</div>
            <div class="metric-value" style="font-size: 1.4rem; color: #818cf8; text-shadow: 0 0 10px rgba(129, 140, 248, 0.4);">{current_total_hours:.2f}h</div>
            <div style="font-size: 0.8rem; color: #cbd5e1; margin-top: 2px;">~{int(current_total_seconds // 60)} phút tập trung</div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  str_app.sidebar.markdown("<br>", unsafe_allow_html=True)

  if str_app.sidebar.button(
      "💾 Lưu lại tiến độ học", type="primary", use_container_width=True
  ):
    if save_current_progress():
      str_app.sidebar.success("✅ Đã lưu tiến độ thành công!")
    else:
      str_app.sidebar.error("❌ Lỗi khi lưu dữ liệu!")

  if str_app.sidebar.button(
      "📧 Gửi Email nhắc nhở báo cáo", use_container_width=True
  ):
    success, err_msg = send_daily_reminder_email(
        SENDER_EMAIL,
        completed_est_count,
        total_all_questions,
        bookmarked_count,
        current_total_hours,
    )
    if success:
      str_app.sidebar.success("✅ Đã gửi email báo cáo vào Gmail!")
    else:
      str_app.sidebar.error(f"❌ Gửi mail thất bại: {err_msg}")

  # NÚT KẾT THÚC ÔN TẬP THÊM MỚI TẠI ĐÂY
  str_app.sidebar.markdown("<br>", unsafe_allow_html=True)
  if str_app.sidebar.button(
      "🏁 Kết thúc phiên ôn tập", type="secondary", use_container_width=True
  ):
    save_current_progress()
    str_app.session_state["app_started"] = False
    scroll_to_top()
    str_app.rerun()

  str_app.sidebar.markdown("<br>", unsafe_allow_html=True)

  str_app.sidebar.markdown(
      """
        <div class="metric-card-container" style="text-align: left;">
            <div class="metric-label" style="margin-bottom: 8px;">🎛️ Điều Hướng Học Tập</div>
    """,
      unsafe_allow_html=True,
  )

  mode = str_app.sidebar.selectbox(
      "Chọn chế độ học:",
      [
          "📖 Ôn tập theo chuyên đề",
          "🔄 Ôn lại câu trả lời sai",
          "🧠 Spaced Repetition (Ôn thông minh)",
          "📂 Ôn gộp tất cả (50 câu/phần)",
          "⭐ Tất cả câu hỏi cần ghi nhớ",
          "📝 Thi thử (Mock Test)",
          "⚙️ Quản lý kho lưu trữ & Dữ liệu",
      ],
      label_visibility="collapsed",
      key="sidebar_mode_select",
  )

  str_app.sidebar.markdown("</div>", unsafe_allow_html=True)

  if error_message:
    str_app.error(error_message)
  else:
    if mode == "📖 Ôn tập theo chuyên đề":
      str_app.sidebar.markdown(
          """
                <div class="metric-card-container" style="text-align: left;">
                    <div class="metric-label" style="margin-bottom: 8px;">📂 Chọn Chuyên Đề</div>
            """,
          unsafe_allow_html=True,
      )
      sheet_list = list(sheets_data.keys())
      selected_sheet = str_app.sidebar.selectbox(
          "Chuyên đề:",
          sheet_list,
          label_visibility="collapsed",
          key="sidebar_chuande_select",
      )
      str_app.sidebar.markdown("</div>", unsafe_allow_html=True)

      q_list = sheets_data[selected_sheet]
      total_q = len(q_list)

      if f"q_idx_{selected_sheet}" not in str_app.session_state:
        str_app.session_state[f"q_idx_{selected_sheet}"] = 0
      if f"done_{selected_sheet}" not in str_app.session_state:
        str_app.session_state[f"done_{selected_sheet}"] = 0

      col_sb1, col_sb2 = str_app.sidebar.columns(2)
      col_sb1.metric("Tổng câu", total_q)
      col_sb2.metric("Đã làm", str_app.session_state[f"done_{selected_sheet}"])

      if str_app.sidebar.button(
          "🔄 Đặt lại chuyên đề này", use_container_width=True
      ):
        str_app.session_state[f"q_idx_{selected_sheet}"] = 0
        str_app.session_state[f"done_{selected_sheet}"] = 0
        keys_to_del = [
            k
            for k in str_app.session_state.keys()
            if k.startswith(f"shuff_chuande_{selected_sheet}_")
            or k.startswith(f"user_ans_chuande_{selected_sheet}_")
            or k.startswith(f"answered_chuande_{selected_sheet}_")
        ]
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

        str_app.markdown(
            f"""
                    <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">📂 Chuyên đề: {selected_sheet}</span>
                            <span class="badge-topic">Câu {idx + 1} / {total_q}</span>
                        </div>
                    </div>
                """,
            unsafe_allow_html=True,
        )

        is_bm = any(
            b.get("question") == q_item["question"]
            for b in str_app.session_state["bookmarked_questions"]
        )
        bm_label = (
            "⭐ Đã đánh dấu ghi nhớ"
            if is_bm
            else "☆ Đánh dấu câu cần ghi nhớ"
        )

        col_q_head1, col_q_head2 = str_app.columns([4, 1])
        with col_q_head2:
          if str_app.button(
              bm_label, key=f"bm_chuande_{idx}", use_container_width=True
          ):
            if is_bm:
              str_app.session_state["bookmarked_questions"] = [
                  b
                  for b in str_app.session_state["bookmarked_questions"]
                  if b.get("question") != q_item["question"]
              ]
              str_app.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
            else:
              str_app.session_state["bookmarked_questions"].append(q_item)
              str_app.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
            save_current_progress()
            str_app.rerun()

        clean_q_text = q_item["question"]
        if clean_q_text.lower().startswith("câu"):
          parts = clean_q_text.split(":", 1)
          if len(parts) > 1:
            clean_q_text = parts[1].strip()

        str_app.markdown(
            f"""
                    <div class="question-card">
                        <div class="question-title">
                            <b>Câu {idx + 1}:</b> {clean_q_text}
                        </div>
                    </div>
                """,
            unsafe_allow_html=True,
        )

        shuff_data = get_shuffled_options(
            q_item, f"shuff_chuande_{selected_sheet}_{idx}"
        )
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
          selected_opt = str_app.radio(
              "Lựa chọn đáp án:",
              options,
              index=default_idx,
              key=f"radio_chuande_{selected_sheet}_{idx}",
              label_visibility="collapsed",
          )

          if selected_opt is not None:
            str_app.session_state[ans_storage_key] = selected_opt
            str_app.session_state[answered_key] = True
            str_app.session_state[f"done_{selected_sheet}"] = min(
                total_q,
                max(str_app.session_state[f"done_{selected_sheet}"], idx + 1),
            )

            is_correct = selected_opt.strip().upper().startswith(correct_letter)
            update_spaced_repetition(q_item["question"], is_correct)
            if not is_correct:
              if not any(
                  w.get("question") == q_item["question"]
                  for w in str_app.session_state["wrong_questions"]
              ):
                str_app.session_state["wrong_questions"].append(q_item)
                save_current_progress()
            str_app.rerun()
        else:
          saved_choice = str_app.session_state.get(ans_storage_key)
          str_app.markdown(
              f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn:"
              f" <b>{saved_choice}</b></p>",
              unsafe_allow_html=True,
          )

          is_correct = saved_choice.strip().upper().startswith(correct_letter)
          str_app.markdown("<br>", unsafe_allow_html=True)
          if is_correct:
            str_app.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
          else:
            correct_text = next(
                (
                    opt
                    for opt in options
                    if opt.strip().upper().startswith(correct_letter)
                ),
                "",
            )
            str_app.error(
                f"❌ Sai rồi! Đáp án đúng là **{correct_text}**."
            )

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
          if str_app.button(
              "Câu tiếp theo ➡️", type="primary", use_container_width=True
          ):
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

        str_app.markdown(
            f"""
                <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 1.05rem; font-weight: 700; color: #f43f5e;">⚠️ Nguồn: {w_item.get('sheet', 'N/A')}</span>
                        <span class="badge-topic">Câu sai {w_idx + 1} / {len(wrong_list)}</span>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        str_app.markdown(
            f"""
                <div class="question-card">
                    <div class="question-title">
                        <b>Câu {w_idx + 1}:</b> {w_item['question']}
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

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
          w_choice = str_app.radio(
              "Lựa chọn đáp án:",
              options,
              index=w_default_idx,
              key=f"radio_wrong_{w_idx}",
              label_visibility="collapsed",
          )
          if w_choice is not None:
            str_app.session_state[w_storage_key] = w_choice
            str_app.session_state[w_answered_key] = True

            is_correct = w_choice.strip().upper().startswith(correct_letter)
            update_spaced_repetition(w_item["question"], is_correct)
            if is_correct:
              str_app.session_state["wrong_questions"] = [
                  w
                  for w in str_app.session_state["wrong_questions"]
                  if w.get("question") != w_item["question"]
              ]
              save_current_progress()
              str_app.toast(
                  "🎉 Đã trả lời đúng! Câu hỏi đã được xóa khỏi danh sách câu"
                  " sai.",
                  icon="✅",
              )
            str_app.rerun()
        else:
          saved_w_choice = str_app.session_state.get(w_storage_key)
          str_app.markdown(
              f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn:"
              f" <b>{saved_w_choice}</b></p>",
              unsafe_allow_html=True,
          )

          is_correct = saved_w_choice.strip().upper().startswith(correct_letter)
          str_app.markdown("<br>", unsafe_allow_html=True)
          if is_correct:
            str_app.success("🎉 Chính xác tuyệt vời!")
          else:
            correct_text = next(
                (
                    opt
                    for opt in options
                    if opt.strip().upper().startswith(correct_letter)
                ),
                "",
            )
            str_app.error(
                f"❌ Chưa chính xác. Đáp án đúng là: **{correct_text}**"
            )

        str_app.markdown("<br>", unsafe_allow_html=True)
        col_w_prev, col_w_next = str_app.columns([1, 1])
        with col_w_prev:
          if str_app.button("⬅️ Câu trước", key="w_prev", use_container_width=True):
            if str_app.session_state["wrong_idx"] > 0:
              str_app.session_state["wrong_idx"] -= 1
            else:
              str_app.session_state["wrong_idx"] = len(wrong_list) - 1
            scroll_to_top()
            str_app.rerun()
        with col_w_next:
          if str_app.button(
              "Câu tiếp theo ➡️",
              key="w_next",
              type="primary",
              use_container_width=True,
          ):
            if str_app.session_state["wrong_idx"] < len(wrong_list) - 1:
              str_app.session_state["wrong_idx"] += 1
            else:
              str_app.session_state["wrong_idx"] = 0
            scroll_to_top()
            str_app.rerun()

    elif mode == "🧠 Spaced Repetition (Ôn thông minh)":
      str_app.title("🧠 Spaced Repetition (Lặp lại ngắt quãng)")
      str_app.markdown("---")
      str_app.info(
          "Chế độ này ưu tiên hiển thị các câu hỏi ông còn yếu dựa trên thuật"
          " toán lặp lại ngắt quãng thông minh."
      )

      all_flat_qs = []
      for s_name, q_list in sheets_data.items():
        all_flat_qs.extend(q_list)

      sr_data = str_app.session_state["spaced_repetition_data"]
      now = time.time()

      due_questions = []
      for q in all_flat_qs:
        q_txt = q["question"]
        if q_txt not in sr_data or sr_data[q_txt]["next_review"] <= now:
          due_questions.append(q)

      if not due_questions:
        str_app.success(
            "🎉 Tuyệt vời! Hiện tại không có câu hỏi nào đến hạn ôn tập thông"
            " minh. Ông có thể ôn tập theo chuyên đề khác nhé!"
        )
      else:
        str_app.write(
            f"Có **{len(due_questions)}** câu hỏi đến lịch ôn tập hôm nay."
        )
        if "sr_idx" not in str_app.session_state:
          str_app.session_state["sr_idx"] = 0

        sr_idx = str_app.session_state["sr_idx"]
        if sr_idx >= len(due_questions):
          sr_idx = 0
          str_app.session_state["sr_idx"] = 0

        sr_item = due_questions[sr_idx]
        box_level = sr_data.get(sr_item["question"], {}).get("box", 1)

        str_app.markdown(
            f"""
                <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 1.05rem; font-weight: 700; color: #818cf8;">🧠 Cấp độ ghi nhớ (Box): Level {box_level} / 5</span>
                        <span class="badge-topic">Câu thông minh {sr_idx + 1} / {len(due_questions)}</span>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        str_app.markdown(
            f"""
                <div class="question-card">
                    <div class="question-title">
                        <b>Câu {sr_idx + 1}:</b> {sr_item['question']}
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        shuff_sr = get_shuffled_options(sr_item, f"shuff_sr_{sr_idx}")
        options = shuff_sr["options"]
        correct_letter = shuff_sr["correct"]

        sr_ans_key = f"user_ans_sr_{sr_idx}"
        sr_ans_checked = f"answered_sr_{sr_idx}"
        is_sr_answered = str_app.session_state.get(sr_ans_checked, False)

        sr_def_idx = None
        if str_app.session_state.get(sr_ans_key) in options:
          sr_def_idx = options.index(str_app.session_state.get(sr_ans_key))

        if not is_sr_answered:
          sr_choice = str_app.radio(
              "Lựa chọn đáp án:",
              options,
              index=sr_def_idx,
              key=f"radio_sr_{sr_idx}",
              label_visibility="collapsed",
          )
          if sr_choice is not None:
            str_app.session_state[sr_ans_key] = sr_choice
            str_app.session_state[sr_ans_checked] = True

            is_correct = sr_choice.strip().upper().startswith(correct_letter)
            update_spaced_repetition(sr_item["question"], is_correct)
            str_app.rerun()
        else:
          saved_sr_choice = str_app.session_state.get(sr_ans_key)
          str_app.markdown(
              f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn:"
              f" <b>{saved_sr_choice}</b></p>",
              unsafe_allow_html=True,
          )

          is_correct = saved_sr_choice.strip().upper().startswith(
              correct_letter
          )
          str_app.markdown("<br>", unsafe_allow_html=True)
          if is_correct:
            str_app.success(
                "🎉 Chính xác! Hộp ghi nhớ đã được nâng cấp lên một bậc."
            )
          else:
            correct_text = next(
                (
                    opt
                    for opt in options
                    if opt.strip().upper().startswith(correct_letter)
                ),
                "",
            )
            str_app.error(
                f"❌ Sai rồi! Đáp án đúng là: **{correct_text}** (Đã reset về"
                " Box 1)"
            )

        str_app.markdown("<br>", unsafe_allow_html=True)
        if str_app.button(
            "Câu tiếp theo ➡️", key="sr_next", type="primary", use_container_width=True
        ):
          if str_app.session_state["sr_idx"] < len(due_questions) - 1:
            str_app.session_state["sr_idx"] += 1
          else:
            str_app.session_state["sr_idx"] = 0
          scroll_to_top()
          str_app.rerun()

    elif mode == "📂 Ôn gộp tất cả (50 câu/phần)":
      str_app.title("📂 Ôn Gộp Tất Cả Câu Hỏi (Chia Chunk 50 câu)")
      str_app.markdown("---")

      all_flat_qs = []
      for s_name, q_list in sheets_data.items():
        all_flat_qs.extend(q_list)

      total_chunks = (len(all_flat_qs) + 49) // 50
      chunk_options = [
          f"Phần {i+1} (Câu {i*50 + 1} đến {min((i+1)*50, len(all_flat_qs))})"
          for i in range(total_chunks)
      ]

      selected_chunk_str = str_app.selectbox(
          "Chọn phần ôn tập:", chunk_options
      )
      chunk_idx = chunk_options.index(selected_chunk_str)

      start_idx = chunk_idx * 50
      end_idx = min((chunk_idx + 1) * 50, len(all_flat_qs))
      chunk_qs = all_flat_qs[start_idx:end_idx]

      str_app.write(
          f"Đang ôn tập phần **{selected_chunk_str}** gồm **{len(chunk_qs)}**"
          " câu hỏi."
      )

      if f"chunk_sub_idx_{chunk_idx}" not in str_app.session_state:
        str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] = 0

      c_idx = str_app.session_state[f"chunk_sub_idx_{chunk_idx}"]
      if c_idx >= len(chunk_qs):
        c_idx = 0
        str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] = 0

      c_item = chunk_qs[c_idx]

      str_app.markdown(
          f"""
            <div class="question-card" style="margin-top: 20px;">
                <div class="question-title">
                    <b>Câu {start_idx + c_idx + 1} (Phần {chunk_idx + 1}):</b> {c_item['question']}
                </div>
            </div>
        """,
          unsafe_allow_html=True,
      )

      shuff_chunk = get_shuffled_options(c_item, f"shuff_chunk_{chunk_idx}_{c_idx}")
      options = shuff_chunk["options"]
      correct_letter = shuff_chunk["correct"]

      chunk_ans_key = f"user_ans_chunk_{chunk_idx}_{c_idx}"
      chunk_answered_key = f"answered_chunk_{chunk_idx}_{c_idx}"
      is_chunk_ans = str_app.session_state.get(chunk_answered_key, False)

      chunk_def_idx = None
      if str_app.session_state.get(chunk_ans_key) in options:
        chunk_def_idx = options.index(str_app.session_state.get(chunk_ans_key))

      if not is_chunk_ans:
        c_choice = str_app.radio(
            "Lựa chọn đáp án:",
            options,
            index=chunk_def_idx,
            key=f"radio_chunk_{chunk_idx}_{c_idx}",
            label_visibility="collapsed",
        )
        if c_choice is not None:
          str_app.session_state[chunk_ans_key] = c_choice
          str_app.session_state[chunk_answered_key] = True

          is_correct = c_choice.strip().upper().startswith(correct_letter)
          update_spaced_repetition(c_item["question"], is_correct)
          if not is_correct:
            if not any(
                w.get("question") == c_item["question"]
                for w in str_app.session_state["wrong_questions"]
            ):
              str_app.session_state["wrong_questions"].append(c_item)
              save_current_progress()
          str_app.rerun()
      else:
        saved_c_choice = str_app.session_state.get(chunk_ans_key)
        str_app.markdown(
            f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn:"
            f" <b>{saved_c_choice}</b></p>",
            unsafe_allow_html=True,
        )

        is_correct = saved_c_choice.strip().upper().startswith(correct_letter)
        str_app.markdown("<br>", unsafe_allow_html=True)
        if is_correct:
          str_app.success("🎉 Chính xác tuyệt vời!")
        else:
          correct_text = next(
              (
                  opt
                  for opt in options
                  if opt.strip().upper().startswith(correct_letter)
              ),
              "",
          )
          str_app.error(f"❌ Sai rồi! Đáp án đúng là: **{correct_text}**")

      str_app.markdown("<br>", unsafe_allow_html=True)
      col_ck1, col_ck2 = str_app.columns([1, 1])
      with col_ck1:
        if str_app.button(
            "⬅️ Câu trước", key=f"ck_prev_{chunk_idx}", use_container_width=True
        ):
          if str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] > 0:
            str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] -= 1
          else:
            str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] = (
                len(chunk_qs) - 1
            )
          scroll_to_top()
          str_app.rerun()
      with col_ck2:
        if str_app.button(
            "Câu tiếp theo ➡️",
            key=f"ck_next_{chunk_idx}",
            type="primary",
            use_container_width=True,
        ):
          if str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] < len(chunk_qs) - 1:
            str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] += 1
          else:
            str_app.session_state[f"chunk_sub_idx_{chunk_idx}"] = 0
            str_app.session_state["completed_chunks"].add(chunk_idx)
            save_current_progress()
            str_app.success(
                f"🎉 Chúc mừng ông đã hoàn thành phần {chunk_idx + 1}!"
            )
          scroll_to_top()
          str_app.rerun()

    elif mode == "⭐ Tất cả câu hỏi cần ghi nhớ":
      str_app.title("⭐ Danh Sách Câu Hỏi Đã Đánh Dấu (Bookmark)")
      str_app.markdown("---")

      bm_list = str_app.session_state["bookmarked_questions"]
      if not bm_list:
        str_app.info(
            "⭐ Hiện chưa có câu hỏi nào được đánh dấu. Hãy nhấn nút sao ở mỗi"
            " câu hỏi để lưu lại nhé!"
        )
      else:
        str_app.write(
            f"Ông đang lưu **{len(bm_list)}** câu hỏi quan trọng cần ghi nhớ."
        )
        for i, bm in enumerate(bm_list):
          with str_app.expander(f"Câu {i+1}: {bm['question'][:80]}..."):
            str_app.markdown(f"**Câu hỏi đầy đủ:** {bm['question']}")
            str_app.markdown(
                f"**Nguồn chuyên đề:** `{bm.get('sheet', 'N/A')}`"
            )
            opts = bm.get("options", [])
            corr = bm.get("correct", "A")
            str_app.markdown("**Các đáp án:**")
            for o in opts:
              is_c = o.strip().upper().startswith(corr)
              if is_c:
                str_app.markdown(
                    f"- ✅ **{o} (Đáp án đúng chuẩn)**",
                    unsafe_allow_html=True,
                )
              else:
                str_app.markdown(f"- {o}")

            if str_app.button(
                "🗑️ Xóa khỏi danh sách ghi nhớ", key=f"del_bm_{i}"
            ):
              str_app.session_state["bookmarked_questions"].pop(i)
              save_current_progress()
              str_app.rerun()

    elif mode == "📝 Thi thử (Mock Test)":
      str_app.title("📝 Kỳ Thi Thử Trực Tuyến (Mock Test 50 Câu)")
      str_app.markdown("---")

      all_flat_qs = []
      for s_name, q_list in sheets_data.items():
        all_flat_qs.extend(q_list)

      if "mock_questions" not in str_app.session_state:
        str_app.session_state["mock_questions"] = random.sample(
            all_flat_qs, min(50, len(all_flat_qs))
        )
        str_app.session_state["mock_submitted"] = False

      mock_qs = str_app.session_state["mock_questions"]

      if not str_app.session_state["mock_submitted"]:
        str_app.write(
            f"Đề thi gồm **{len(mock_qs)}** câu hỏi ngẫu nhiên. Hãy hoàn thành"
            " và bấm nộp bài ở cuối trang."
        )

        mock_user_answers = {}
        for idx, mq in enumerate(mock_qs):
          str_app.markdown(
              f"**Câu {idx+1}:** {mq['question']}"
          )
          shuff_m = get_shuffled_options(mq, f"shuff_mock_{idx}")
          m_opts = shuff_m["options"]
          mock_user_answers[idx] = str_app.radio(
              "Chọn đáp án:",
              m_opts,
              key=f"radio_mock_ans_{idx}",
              label_visibility="collapsed",
          )
          str_app.markdown("---")

        if str_app.button(
            "📤 Nộp bài thi thử", type="primary", use_container_width=True
        ):
          score = 0
          for idx, mq in enumerate(mock_qs):
            shuff_m = str_app.session_state[f"shuff_mock_{idx}"]
            corr_l = shuff_m["correct"]
            chosen = mock_user_answers.get(idx, "")
            if chosen and chosen.strip().upper().startswith(corr_l):
              score += 1

          str_app.session_state["mock_score"] = score
          str_app.session_state["mock_submitted"] = True
          if score >= 40:
            str_app.session_state["passed_tests"].add("mock_test")
          save_current_progress()
          scroll_to_top()
          str_app.rerun()
      else:
        score = str_app.session_state.get("mock_score", 0)
        total_m = len(mock_qs)
        pct = (score / total_m) * 100

        str_app.markdown(
            f"""
                <div class="welcome-card" style="margin-bottom: 30px;">
                    <h2>🎯 KẾT QUẢ BÀI THI THỬ</h2>
                    <div class="metric-value" style="font-size: 2.5rem; margin: 15px 0;">{score} / {total_m} câu đúng</div>
                    <p style="font-size: 1.2rem; color: #38bdf8;">Tỉ lệ chính xác: {pct:.1f}%</p>
                </div>
            """,
            unsafe_allow_html=True,
        )

        if pct >= 80:
          str_app.success(
              "🎉 Xuất sắc! Kiến thức an toàn điện của ông đã rất vững vàng!"
          )
        else:
          str_app.warning(
              "⚠️ Cần cố gắng ôn tập thêm các câu trả lời sai để đạt kết quả tốt"
              " hơn nhé!"
          )

        if str_app.button("🔄 Làm đề thi thử mới", type="primary"):
          if "mock_questions" in str_app.session_state:
            del str_app.session_state["mock_questions"]
          keys_to_del = [
              k
              for k in str_app.session_state.keys()
              if k.startswith("shuff_mock_") or k.startswith("radio_mock_ans_")
          ]
          for k in keys_to_del:
            del str_app.session_state[k]
          str_app.session_state["mock_submitted"] = False
          scroll_to_top()
          str_app.rerun()

    elif mode == "⚙️ Quản lý kho lưu trữ & Dữ liệu":
      str_app.title("⚙️ Quản Lý Kho Lưu Trữ & Dữ Liệu")
      str_app.markdown("---")

      str_app.subheader("📊 Thống kê chi tiết")
      str_app.write(f"- Tổng số câu hỏi trong ngân hàng: **{total_all_questions}** câu")
      str_app.write(f"- Số câu hỏi đã đánh dấu (Star): **{len(str_app.session_state['bookmarked_questions'])}** câu")
      str_app.write(f"- Số câu trả lời sai đang lưu: **{len(str_app.session_state['wrong_questions'])}** câu")
      str_app.write(f"- Tổng thời gian ôn tập tích lũy: **{current_total_hours:.2f} giờ**")

      str_app.markdown("---")
      str_app.subheader("🛠️ Thao tác dữ liệu")

      col_d1, col_d2 = str_app.columns(2)
      with col_d1:
        if str_app.button("🗑️ Xóa toàn bộ danh sách câu sai", use_container_width=True):
          str_app.session_state["wrong_questions"] = []
          save_current_progress()
          str_app.success("Đã xóa danh sách câu sai!")
          str_app.rerun()

      with col_d2:
        if str_app.button("🔄 Reset toàn bộ tiến độ học tập", type="primary", use_container_width=True):
          if os.path.exists(PROGRESS_FILE):
            os.remove(PROGRESS_FILE)
          str_app.session_state["wrong_questions"] = []
          str_app.session_state["bookmarked_questions"] = []
          str_app.session_state["completed_chunks"] = set()
          str_app.session_state["passed_tests"] = set()
          str_app.session_state["spaced_repetition_data"] = {}
          str_app.success("Đã reset toàn bộ tiến độ thành công!")
          str_app.rerun()
