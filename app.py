import streamlit as st
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

st.set_page_config(
    page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện",
    page_icon="⚡",
    layout="wide"
)

PROGRESS_FILE = "quiz_progress.json"

# --- CẤU HÌNH GMAIL GỬI THÔNG BÁO ---
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
        "total_study_seconds": 0,
        "last_login_date": ""
    }

def save_current_progress():
    if "start_session_time" in st.session_state:
        elapsed = time.time() - st.session_state["start_session_time"]
        st.session_state["start_session_time"] = time.time()
        saved_data = load_saved_progress()
        saved_data["total_study_seconds"] = saved_data.get("total_study_seconds", 0) + elapsed
    else:
        saved_data = load_saved_progress()

    data = {
        "completed_chunks": list(st.session_state.get("completed_chunks", [])),
        "passed_tests": list(st.session_state.get("passed_tests", [])),
        "bookmarked_questions": st.session_state.get("bookmarked_questions", []),
        "wrong_questions": st.session_state.get("wrong_questions", []),
        "total_study_seconds": saved_data.get("total_study_seconds", 0),
        "last_login_date": st.session_state.get("last_login_date", "")
    }
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except:
        return False

def send_daily_reminder_email(receiver_email, completed_questions_count, total_questions_count, bookmarked_count, total_study_hours):
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
    message["Subject"] = Header(subject, 'utf-8')
    message.attach(MIMEText(body, 'plain', 'utf-8'))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, receiver_email, message.as_string())
        server.quit()
        return True, "Thành công"
    except Exception as e:
        return False, str(e)

@st.cache_data
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

if "start_session_time" not in st.session_state:
    st.session_state["start_session_time"] = time.time()

sheets_data, error_message = load_data()
saved_prog = load_saved_progress()

total_all_questions = 0
if sheets_data:
    for sh, ql in sheets_data.items():
        total_all_questions += len(ql)

if "wrong_questions" not in st.session_state:
    st.session_state["wrong_questions"] = saved_prog.get("wrong_questions", [])
if "bookmarked_questions" not in st.session_state:
    st.session_state["bookmarked_questions"] = saved_prog.get("bookmarked_questions", [])
if "completed_chunks" not in st.session_state:
    st.session_state["completed_chunks"] = set(saved_prog.get("completed_chunks", []))
if "passed_tests" not in st.session_state:
    st.session_state["passed_tests"] = set(saved_prog.get("passed_tests", []))
if "app_started" not in st.session_state:
    st.session_state["app_started"] = False

save_current_progress()

def scheduled_job():
    saved_data = load_saved_progress()
    bm_cnt = len(saved_data.get("bookmarked_questions", []))
    completed_chunks_set = set(saved_data.get("completed_chunks", [])).union(set(saved_data.get("passed_tests", [])))
    comp_q_cnt = min(len(completed_chunks_set) * 50, total_all_questions)
    total_hours = saved_data.get("total_study_seconds", 0) / 3600.0
    send_daily_reminder_email(SENDER_EMAIL, comp_q_cnt, total_all_questions, bm_cnt, total_hours)

@st.cache_resource
def start_scheduler():
    scheduler = BackgroundScheduler()
    scheduler.add_job(scheduled_job, 'cron', hour=7, minute=0)
    scheduler.start()
    return scheduler

try:
    start_scheduler()
except Exception:
    pass

st.markdown("""
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
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.7), rgba(15, 23, 42, 0.9));
        border: 1px solid rgba(56, 189, 248, 0.3);
        box-shadow: 0 0 25px rgba(56, 189, 248, 0.15);
        border-radius: 20px;
        padding: 40px;
        text-align: center;
        backdrop-filter: blur(10px);
    }
    .sidebar-card {
        background: rgba(30, 41, 59, 0.5);
        border: 1px solid rgba(56, 189, 248, 0.2);
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 15px;
    }
    
    .main-header-card {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.6), rgba(15, 23, 42, 0.8));
        border-left: 5px solid #38bdf8;
        padding: 18px 22px;
        border-radius: 0 12px 12px 0;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .question-card {
        background: rgba(30, 41, 59, 0.4);
        border: 1px solid rgba(56, 189, 248, 0.25);
        border-radius: 16px;
        padding: 28px;
        margin-bottom: 25px;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.25);
        backdrop-filter: blur(8px);
    }
    .question-title {
        font-size: 1.25rem !important;
        line-height: 1.6;
        color: #f8fafc !important;
        font-weight: 600;
        margin-bottom: 15px;
    }
    .badge-topic {
        background-color: rgba(56, 189, 248, 0.15);
        color: #38bdf8;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        border: 1px solid rgba(56, 189, 248, 0.3);
    }
    
    div[data-testid="stRadio"] > div[role="radiogroup"] label p,
    div[data-testid="stRadio"] > div[role="radiogroup"] label span,
    div[data-testid="stRadio"] > div[role="radiogroup"] label div {
        font-size: 1.05rem !important;
        line-height: 1.5 !important;
        color: #f1f5f9 !important;
    }
    </style>
""", unsafe_allow_html=True)

def get_shuffled_options(q_item, session_key):
    if session_key not in st.session_state:
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
                
        st.session_state[session_key] = {
            "options": shuffled_options,
            "correct": new_correct_letter
        }
        
    return st.session_state[session_key]

if not st.session_state["app_started"]:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_w1, col_w2, col_w3 = st.columns([1, 2.2, 1])
    with col_w2:
        st.markdown("""
            <div class="welcome-card">
                <h1 class="sparkle-title">⚡ Chào Đức Anh!</h1>
                <p style="font-size: 1.25rem; color: #e2e8f0; margin-top: 15px; font-weight: 500;">Hệ thống Ngân hàng câu hỏi An Toàn Điện đã sẵn sàng.</p>
                <p style="font-size: 1.05rem; color: #94a3b8; margin-top: 8px;">Chúc ông ôn tập thật tập trung, nắm trọn kiến thức và đạt kết quả cao nhất! 🚀</p>
            </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        col_btn_center1, col_btn_center2, col_btn_center3 = st.columns([1, 1.5, 1])
        with col_btn_center2:
            if st.button("✨ Bắt đầu vào ôn tập ngay", type="primary", use_container_width=True):
                st.session_state["app_started"] = True
                st.rerun()
else:
    st.sidebar.markdown("### ⚡ Dashboard Ôn Tập")

    completed_chunks_set = st.session_state["completed_chunks"].union(st.session_state["passed_tests"])
    chunk_size_calc = 50
    completed_est_count = min(len(completed_chunks_set) * chunk_size_calc, total_all_questions)
    progress_ratio = completed_est_count / total_all_questions if total_all_questions > 0 else 0.0

    st.sidebar.markdown(f"""
        <div class="sidebar-card">
            <span style="font-size: 0.9rem; font-weight: 600; color: #38bdf8;">📈 TIẾN ĐỘ TỔNG QUAN</span>
        </div>
    """, unsafe_allow_html=True)
    st.sidebar.progress(progress_ratio)
    st.sidebar.caption(f"Đã hoàn thành: **{completed_est_count}/{total_all_questions}** câu ({progress_ratio * 100:.1f}%)")

    current_total_seconds = saved_prog.get("total_study_seconds", 0) + (time.time() - st.session_state["start_session_time"])
    current_total_hours = current_total_seconds / 3600.0

    st.sidebar.info(f"⏱️ Tổng thời gian: **{current_total_hours:.2f}h** (~{int(current_total_seconds // 60)} phút)")

    if st.sidebar.button("💾 Lưu lại tiến độ học", type="primary", use_container_width=True):
        if save_current_progress():
            st.sidebar.success("✅ Đã lưu tiến độ thành công!")
        else:
            st.sidebar.error("❌ Lỗi khi lưu dữ liệu!")

    if st.sidebar.button("📧 Gửi Email nhắc nhở báo cáo", use_container_width=True):
        bm_cnt = len(st.session_state["bookmarked_questions"])
        success, err_msg = send_daily_reminder_email(SENDER_EMAIL, completed_est_count, total_all_questions, bm_cnt, current_total_hours)
        if success:
            st.sidebar.success("✅ Đã gửi email báo cáo vào Gmail!")
        else:
            st.sidebar.error(f"❌ Gửi mail thất bại: {err_msg}")

    st.sidebar.markdown("---")
    
    st.sidebar.markdown("### 🎛️ Điều Hướng Học Tập")
    mode = st.sidebar.selectbox("Chọn chế độ học:", [
        "📖 Ôn tập theo chuyên đề",
        "🔄 Ôn lại câu trả lời sai",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "⭐ Tất cả câu hỏi cần ghi nhớ",
        "📝 Thi thử (Mock Test)"
    ], label_visibility="collapsed")

    if error_message:
        st.error(error_message)
    else:
        if mode == "📖 Ôn tập theo chuyên đề":
            st.sidebar.markdown("---")
            st.sidebar.markdown("#### 📂 Chọn Chuyên Đề")
            sheet_list = list(sheets_data.keys())
            selected_sheet = st.sidebar.selectbox("Chuyên đề:", sheet_list, label_visibility="collapsed")
            
            q_list = sheets_data[selected_sheet]
            total_q = len(q_list)
            
            if f"q_idx_{selected_sheet}" not in st.session_state:
                st.session_state[f"q_idx_{selected_sheet}"] = 0
            if f"done_{selected_sheet}" not in st.session_state:
                st.session_state[f"done_{selected_sheet}"] = 0
                
            st.sidebar.markdown("---")
            col_sb1, col_sb2 = st.sidebar.columns(2)
            col_sb1.metric("Tổng câu", total_q)
            col_sb2.metric("Đã làm", st.session_state[f'done_{selected_sheet}'])
            
            if st.sidebar.button("🔄 Đặt lại chuyên đề này", use_container_width=True):
                st.session_state[f"q_idx_{selected_sheet}"] = 0
                st.session_state[f"done_{selected_sheet}"] = 0
                keys_to_del = [k for k in st.session_state.keys() if k.startswith(f"shuff_chuande_{selected_sheet}_") or k.startswith(f"user_ans_chuande_{selected_sheet}_")]
                for k in keys_to_del:
                    del st.session_state[k]
                st.rerun()

            st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
            st.markdown("---")
            
            idx = st.session_state[f"q_idx_{selected_sheet}"]
            if idx >= total_q and total_q > 0:
                idx = 0
                st.session_state[f"q_idx_{selected_sheet}"] = 0
                
            if total_q == 0:
                st.warning("Chuyên đề này hiện chưa có câu hỏi nào.")
            else:
                q_item = q_list[idx]
                
                st.markdown(f"""
                    <div class="main-header-card">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.1rem; font-weight: 700; color: #38bdf8;">📂 Chuyên đề: {selected_sheet}</span>
                            <span class="badge-topic">Câu {idx + 1} / {total_q}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<div class="question-card">', unsafe_allow_html=True)
                
                is_bm = any(b.get("question") == q_item["question"] for b in st.session_state["bookmarked_questions"])
                bm_label = "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
                
                col_q_head1, col_q_head2 = st.columns([4, 1])
                with col_q_head2:
                    if st.button(bm_label, key=f"bm_chuande_{idx}", use_container_width=True):
                        if is_bm:
                            st.session_state["bookmarked_questions"] = [b for b in st.session_state["bookmarked_questions"] if b.get("question") != q_item["question"]]
                            st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
                        else:
                            st.session_state["bookmarked_questions"].append(q_item)
                            st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
                        save_current_progress()
                        st.rerun()
                
                clean_q_text = q_item['question']
                if clean_q_text.lower().startswith("câu"):
                    parts = clean_q_text.split(":", 1)
                    if len(parts) > 1:
                        clean_q_text = parts[1].strip()

                st.markdown(f"""
                    <div class="question-title">
                        <b>Câu {idx + 1}:</b> {clean_q_text}
                    </div>
                """, unsafe_allow_html=True)
                
                shuff_data = get_shuffled_options(q_item, f"shuff_chuande_{selected_sheet}_{idx}")
                options = shuff_data["options"]
                correct_letter = shuff_data["correct"]
                
                ans_storage_key = f"user_ans_chuande_{selected_sheet}_{idx}"
                default_idx = None
                current_saved_ans = st.session_state.get(ans_storage_key, None)
                if current_saved_ans in options:
                    default_idx = options.index(current_saved_ans)
                    
                selected_opt = st.radio("Chọn đáp án của bạn:", options, index=default_idx, key=f"radio_chuande_{selected_sheet}_{idx}")
                
                if selected_opt is not None:
                    st.session_state[ans_storage_key] = selected_opt
                    is_correct = selected_opt.strip().upper().startswith(correct_letter)
                    st.markdown("<br>", unsafe_allow_html=True)
                    if is_correct:
                        st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                    else:
                        st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
                        if not any(w.get("question") == q_item["question"] for w in st.session_state["wrong_questions"]):
                            st.session_state["wrong_questions"].append(q_item)
                            save_current_progress()
                            
                    st.session_state[f"done_{selected_sheet}"] = min(total_q, max(st.session_state[f"done_{selected_sheet}"], idx + 1))
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                col_prev, col_next = st.columns([1, 1])
                with col_prev:
                    if st.button("⬅️ Câu trước", use_container_width=True):
                        if st.session_state[f"q_idx_{selected_sheet}"] > 0:
                            st.session_state[f"q_idx_{selected_sheet}"] -= 1
                        else:
                            st.session_state[f"q_idx_{selected_sheet}"] = total_q - 1
                        st.rerun()
                with col_next:
                    if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True):
                        if st.session_state[f"q_idx_{selected_sheet}"] < total_q - 1:
                            st.session_state[f"q_idx_{selected_sheet}"] += 1
                        else:
                            st.session_state[f"q_idx_{selected_sheet}"] = 0
                        st.rerun()

        elif mode == "🔄 Ôn lại câu trả lời sai":
            st.title("🔄 Ôn Lại Câu Trả Lời Sai")
            st.markdown("---")
            
            wrong_list = st.session_state["wrong_questions"]
            if not wrong_list:
                st.info("🎉 Hiện tại bạn chưa có câu trả lời sai nào được lưu lại.")
            else:
                st.write(f"Bạn đang có **{len(wrong_list)}** câu cần ôn tập lại.")
                if "wrong_idx" not in st.session_state:
                    st.session_state["wrong_idx"] = 0
                    
                w_idx = st.session_state["wrong_idx"]
                if w_idx >= len(wrong_list):
                    w_idx = 0
                    st.session_state["wrong_idx"] = 0
                    
                w_item = wrong_list[w_idx]
                
                st.markdown(f"""
                    <div class="main-header-card">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.1rem; font-weight: 700; color: #f43f5e;">⚠️ Nguồn: {w_item.get('sheet', 'N/A')}</span>
                            <span class="badge-topic">Câu sai {w_idx + 1} / {len(wrong_list)}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<div class="question-card">', unsafe_allow_html=True)
                st.markdown(f"""
                    <div class="question-title">
                        <b>Câu {w_idx + 1}:</b> {w_item['question']}
                    </div>
                """, unsafe_allow_html=True)
                
                shuff_data = get_shuffled_options(w_item, f"shuff_wrong_{w_idx}")
                options = shuff_data["options"]
                correct_letter = shuff_data["correct"]
                
                w_storage_key = f"user_ans_wrong_{w_idx}"
                w_default_idx = None
                if st.session_state.get(w_storage_key) in options:
                    w_default_idx = options.index(st.session_state.get(w_storage_key))
                    
                w_choice = st.radio("Chọn đáp án của bạn:", options, index=w_default_idx, key=f"radio_wrong_{w_idx}")
                if w_choice is not None:
                    st.session_state[w_storage_key] = w_choice
                    st.markdown("<br>", unsafe_allow_html=True)
                    if w_choice.strip().upper().startswith(correct_letter):
                        st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                        wrong_list = [w for w in wrong_list if w.get("question") != w_item.get("question")]
                        st.session_state["wrong_questions"] = wrong_list
                        save_current_progress()
                    else:
                        st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
                st.markdown('</div>', unsafe_allow_html=True)
                
                if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_wrong_{w_idx}"):
                    if st.session_state["wrong_idx"] < len(wrong_list) - 1:
                        st.session_state["wrong_idx"] += 1
                    else:
                        st.session_state["wrong_idx"] = 0
                    st.rerun()

        elif mode == "⭐ Tất cả câu hỏi cần ghi nhớ":
            st.title("⭐ Tất Cả Câu Hỏi Cần Ghi Nhớ")
            st.markdown("---")
            
            bm_list = st.session_state["bookmarked_questions"]
            if not bm_list:
                st.info("⭐ Bạn chưa đánh dấu câu hỏi nào cần ghi nhớ cả.")
            else:
                st.write(f"Tổng số câu bạn đã đánh dấu ghi nhớ: **{len(bm_list)}** câu.")
                if "global_bm_idx" not in st.session_state:
                    st.session_state["global_bm_idx"] = 0
                    
                gbm_idx = st.session_state["global_bm_idx"]
                if gbm_idx >= len(bm_list):
                    gbm_idx = 0
                    st.session_state["global_bm_idx"] = 0
                    
                bm_item = bm_list[gbm_idx]
                
                st.markdown(f"""
                    <div class="main-header-card">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <span style="font-size: 1.1rem; font-weight: 700; color: #eab308;">⭐ Chuyên đề: {bm_item.get('sheet', 'N/A')}</span>
                            <span class="badge-topic">Đã lưu {gbm_idx + 1} / {len(bm_list)}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown('<div class="question-card">', unsafe_allow_html=True)
                
                col_bm1, col_bm2 = st.columns([4, 1])
                with col_bm2:
                    if st.button("❌ Bỏ lưu", key=f"remove_global_bm_{gbm_idx}__", use_container_width=True):
                        st.session_state["bookmarked_questions"] = [b for b in st.session_state["bookmarked_questions"] if b.get("question") != bm_item.get("question")]
                        save_current_progress()
                        st.toast("Đã xóa khỏi danh sách ghi nhớ!", icon="ℹ️")
                        st.rerun()

                st.markdown(f"""
                    <div class="question-title">
                        <b>Câu hỏi:</b> {bm_item['question']}
                    </div>
                """, unsafe_allow_html=True)
                
                options = bm_item["options"]
                if not options:
                    options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
                
                correct_letter = bm_item.get("correct", "A").strip().upper()
                
                st.markdown("##### 💡 Đáp án chuẩn:")
                for opt in options:
                    opt_letter = opt.strip().upper()[:1]
                    if opt_letter == correct_letter:
                        st.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
                    else:
                        st.markdown(f"- {opt}")
                st.markdown('</div>', unsafe_allow_html=True)
                
                if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_global_bm_{gbm_idx}"):
                    if st.session_state["global_bm_idx"] < len(bm_list) - 1:
                        st.session_state["global_bm_idx"] += 1
                    else:
                        st.session_state["global_bm_idx"] = 0
                    st.rerun()

        elif mode == "📂 Ôn gộp tất cả (50 câu/phần)":
            st.title("📂 Ôn Gộp Tất Cả Chuyên Đề (50 Câu/Phần)")
            st.markdown("---")
            
            all_questions = []
            for sh, ql in sheets_data.items():
                all_questions.extend(ql)
                
            total_all = len(all_questions)
            if total_all == 0:
                st.warning("Không có dữ liệu câu hỏi.")
            else:
                chunk_size = 50
                total_chunks = (total_all // chunk_size) + (1 if total_all % chunk_size != 0 else 0)
                
                st.sidebar.markdown("---")
                st.sidebar.markdown("#### 📂 Chọn Phần Ôn Tập")
                
                chunk_names = []
                for i in range(total_chunks):
                    is_done = (i in st.session_state["completed_chunks"] or i in st.session_state["passed_tests"])
                    prefix = "✅ " if is_done else "📌 "
                    chunk_names.append(f"{prefix}Phần {i+1}")
                    
                selected_chunk_name = st.sidebar.selectbox("Chọn phần:", chunk_names, label_visibility="collapsed")
                c_num = int(selected_chunk_name.split("Phần")[1].strip()) - 1
                
                start_idx = c_num * chunk_size
                end_idx = min((c_num + 1) * chunk_size, total_all)
                
                if f"chunk_q_{c_num}" not in st.session_state:
                    chunk_qs = all_questions.copy()
                    random.seed(42)
                    random.shuffle(chunk_qs)
                    st.session_state[f"chunk_q_{c_num}"] = chunk_qs[start_idx:end_idx]
                    
                current_chunk_questions = st.session_state[f"chunk_q_{c_num}"]
                actual_chunk_len = len(current_chunk_questions)
                
                st.sidebar.markdown("---")
                st.sidebar.metric("Tổng số câu của phần", actual_chunk_len)
                
                sub_mode = st.radio("Chế độ học trong phần:", [
                    "📖 Ôn tập từng câu",
                    "📝 Bài kiểm tra chốt kiến thức",
                    "🔄 Làm lại phần này",
                    "⚠️ Ôn các câu sai"
                ], horizontal=True, label_visibility="collapsed")
                st.markdown("---")
                
                if sub_mode == "📖 Ôn tập từng câu":
                    if f"gop_idx_{c_num}" not in st.session_state:
                        st.session_state[f"gop_idx_{c_num}"] = 0
                    
                    g_idx = st.session_state[f"gop_idx_{c_num}"]
                    
                    if g_idx >= actual_chunk_len:
                        st.success(f"🎉 Bạn đã hoàn thành phần ôn tập (Phần {c_num + 1})!")
                        st.markdown("---")
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            if st.button("🔄 Ôn lại từ đầu", type="secondary", use_container_width=True):
                                st.session_state[f"gop_idx_{c_num}"] = 0
                                keys_to_del = [k for k in st.session_state.keys() if k.startswith(f"shuff_gop_{c_num}_") or k.startswith(f"user_ans_gop_{c_num}_")]
                                for k in keys_to_del:
                                    del st.session_state[k]
                                st.rerun()
                        with col_btn2:
                            if st.button("📝 Bắt đầu bài kiểm tra chốt", type="primary", use_container_width=True):
                                st.session_state[f"auto_switch_test_{c_num}"] = True
                                st.rerun()
                    else:
                        if st.session_state.get(f"auto_switch_test_{c_num}", False):
                            st.session_state[f"auto_switch_test_{c_num}"] = False
                            sub_mode = "📝 Bài kiểm tra chốt kiến thức"

                    if sub_mode == "📖 Ôn tập từng câu" and g_idx < actual_chunk_len:
                        q_item = current_chunk_questions[g_idx]
                        
                        st.markdown(f"""
                            <div class="main-header-card">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 1.1rem; font-weight: 700; color: #38bdf8;">📂 Phần {c_num + 1} — Chuyên đề: {q_item['sheet']}</span>
                                    <span class="badge-topic">Câu {g_idx + 1} / {actual_chunk_len}</span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown('<div class="question-card">', unsafe_allow_html=True)
                        
                        is_bm = any(b.get("question") == q_item["question"] for b in st.session_state["bookmarked_questions"])
                        bm_label = "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
                        
                        col_q_head1, col_q_head2 = st.columns([4, 1])
                        with col_q_head2:
                            if st.button(bm_label, key=f"bm_gop_{c_num}_{g_idx}", use_container_width=True):
                                if is_bm:
                                    st.session_state["bookmarked_questions"] = [b for b in st.session_state["bookmarked_questions"] if b.get("question") != q_item["question"]]
                                    st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
                                else:
                                    st.session_state["bookmarked_questions"].append(q_item)
                                    st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
                                save_current_progress()
                                st.rerun()
                                
                        st.markdown(f"""
                            <div class="question-title">
                                <b>Câu {g_idx + 1}:</b> {q_item['question']}
                            </div>
                        """, unsafe_allow_html=True)
                        
                        shuff_data = get_shuffled_options(q_item, f"shuff_gop_{c_num}_{g_idx}")
                        options = shuff_data["options"]
                        correct_letter = shuff_data["correct"]
                        
                        gop_storage_key = f"user_ans_gop_{c_num}_{g_idx}"
                        gop_default_idx = None
                        if st.session_state.get(gop_storage_key) in options:
                            gop_default_idx = options.index(st.session_state.get(gop_storage_key))
                            
                        g_choice = st.radio("Chọn đáp án của bạn:", options, index=gop_default_idx, key=f"radio_gop_{c_num}_{g_idx}")
                        
                        if g_choice is not None:
                            st.session_state[gop_storage_key] = g_choice
                            is_correct = g_choice.strip().upper().startswith(correct_letter)
                            st.markdown("<br>", unsafe_allow_html=True)
                            if is_correct:
                                st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                            else:
                                st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
                                if not any(w.get("question") == q_item["question"] for w in st.session_state["wrong_questions"]):
                                    st.session_state["wrong_questions"].append(q_item)
                                    save_current_progress()
                                    
                        st.markdown('</div>', unsafe_allow_html=True)
                        
                        col_prev, col_next = st.columns(2)
                        with col_prev:
                            if st.button("⬅️ Câu trước", use_container_width=True):
                                if st.session_state[f"gop_idx_{c_num}"] > 0:
                                    st.session_state[f"gop_idx_{c_num}"] -= 1
                                else:
                                    st.session_state[f"gop_idx_{c_num}"] = actual_chunk_len - 1
                                st.rerun()
                        with col_next:
                            if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True):
                                st.session_state[f"gop_idx_{c_num}"] += 1
                                st.rerun()

                elif sub_mode == "📝 Bài kiểm tra chốt kiến thức":
                    st.markdown(f"### Bài Kiểm Tra - Phần {c_num + 1} ({actual_chunk_len} câu)")
                    st.markdown("---")
                    
                    submitted_key = f"submitted_test_{c_num}"
                    if submitted_key not in st.session_state:
                        st.session_state[submitted_key] = False
                        
                    if not st.session_state[submitted_key]:
                        user_answers = {}
                        for i, q in enumerate(current_chunk_questions):
                            st.markdown('<div class="question-card">', unsafe_allow_html=True)
                            st.markdown(f"""
                                <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                                    <span class="badge-topic">{q['sheet']}</span>
                                    <span style="color: #94a3b8; font-weight: 600;">Câu {i+1}</span>
                                </div>
                                <div class="question-title" style="font-size: 1.1rem !important;">
                                    {q['question']}
                                </div>
                            """, unsafe_allow_html=True)
                            
                            shuff_data = get_shuffled_options(q, f"shuff_test_{c_num}_{i}")
                            options = shuff_data["options"]
                            
                            test_ans_key = f"test_chunk_{c_num}_{i}"
                            ans = st.radio("Chọn đáp án:", options, index=(options.index(st.session_state[test_ans_key]) if st.session_state.get(test_ans_key) in options else None), key=test_ans_key, label_visibility="collapsed")
                            user_answers[i] = ans
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                        if st.button("📤 Nộp bài kiểm tra ngay", type="primary", use_container_width=True):
                            unanswered = [i + 1 for i in range(actual_chunk_len) if user_answers[i] is None]
                            if unanswered:
                                st.error(f"⚠️ Ông chưa làm xong tất cả các câu! Còn thiếu các câu: {', '.join(map(str, unanswered))}.")
                            else:
                                correct_count = 0
                                wrong_count = 0
                                for i, q in enumerate(current_chunk_questions):
                                    selected = user_answers[i]
                                    shuff_data = st.session_state.get(f"shuff_test_{c_num}_{i}", {"correct": "A"})
                                    correct_letter = shuff_data["correct"]
                                    
                                    if selected and selected.strip().upper().startswith(correct_letter):
                                        correct_count += 1
                                    else:
                                        wrong_count += 1
                                        if not any(w.get("question") == q["question"] for w in st.session_state["wrong_questions"]):
                                            st.session_state["wrong_questions"].append(q)
                                            
                                st.session_state[f"result_correct_{c_num}"] = correct_count
                                st.session_state[f"result_wrong_{c_num}"] = wrong_count
                                st.session_state[f"test_user_answers_{c_num}"] = user_answers
                                st.session_state[submitted_key] = True
                                st.session_state["completed_chunks"].add(c_num)
                                st.session_state["passed_tests"].add(c_num)
                                save_current_progress()
                                st.balloons()
                                st.rerun()
                    else:
                        c_correct = st.session_state.get(f"result_correct_{c_num}", 0)
                        c_wrong = st.session_state.get(f"result_wrong_{c_num}", 0)
                        score_percent = (c_correct / actual_chunk_len) * 100
                        
                        st.success("🎉 Đã nộp bài kiểm tra thành công!")
                        st.markdown("### 📊 Kết Quả Bài Kiểm Tra")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Số câu đúng", f"{c_correct}/{actual_chunk_len}", f"{score_percent:.1f}%")
                        col2.metric("Số câu sai", f"{c_wrong}/{actual_chunk_len}")
                        col3.metric("Trạng thái", "Đã hoàn thành ✅")
                        
                        st.markdown("---")
                        st.markdown("### 🔍 Xem Lại Chi Tiết Các Câu Trả Lời Sai")
                        user_answers = st.session_state.get(f"test_user_answers_{c_num}", {})
                        wrong_items_in_test = []
                        for i, q in enumerate(current_chunk_questions):
                            selected = user_answers.get(i)
                            shuff_data = st.session_state.get(f"shuff_test_{c_num}_{i}", {"correct": "A"})
                            correct_letter = shuff_data["correct"]
                            
                            is_correct = selected and selected.strip().upper().startswith(correct_letter)
                            if not is_correct:
                                wrong_items_in_test.append((i, q, selected, shuff_data))
                                
                        if not wrong_items_in_test:
                            st.info("🎉 Tuyệt vời! Ông đã trả lời đúng tất cả các câu trong phần này.")
                        else:
                            for q_idx, q_item, user_sel, shuff_info in wrong_items_in_test:
                                st.markdown('<div class="question-card">', unsafe_allow_html=True)
                                st.markdown(f"**Câu {q_idx + 1}** *(Thuộc chuyên đề: {q_item['sheet']})*")
                                st.markdown(f"> **{q_item['question']}**")
                                correct_letter = shuff_info["correct"]
                                for opt in shuff_info["options"]:
                                    opt_letter = opt.strip().upper()[:1]
                                    is_this_correct = (opt_letter == correct_letter)
                                    is_user_chosen = (user_sel and opt.strip() == user_sel.strip())
                                    
                                    if is_this_correct:
                                        st.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
                                    elif is_user_chosen:
                                        st.markdown(f"- ❌ ~~{opt}~~ *(Ông đã chọn)*")
                                    else:
                                        st.markdown(f"- {opt}")
                                st.markdown('</div>', unsafe_allow_html=True)
                                
                        if st.button("🔄 Làm lại bài kiểm tra này", type="primary", use_container_width=True):
                            st.session_state[submitted_key] = False
                            keys_to_del = [k for k in st.session_state.keys() if k.startswith(f"shuff_test_{c_num}_") or k.startswith(f"test_chunk_{c_num}_")]
                            for k in keys_to_del:
                                del st.session_state[k]
                            if f"test_user_answers_{c_num}" in st.session_state:
                                del st.session_state[f"test_user_answers_{c_num}"]
                            if c_num in st.session_state["completed_chunks"]:
                                st.session_state["completed_chunks"].remove(c_num)
                            if c_num in st.session_state["passed_tests"]:
                                st.session_state["passed_tests"].remove(c_num)
                            save_current_progress()
                            st.rerun()

                elif sub_mode == "🔄 Làm lại phần này":
                    st.session_state[f"gop_idx_{c_num}"] = 0
                    keys_to_del = [k for k in st.session_state.keys() if k.startswith(f"shuff_gop_{c_num}_") or k.startswith(f"shuff_test_{c_num}_") or k.startswith(f"user_ans_gop_{c_num}_") or k.startswith(f"test_chunk_{c_num}_")]
                    for k in keys_to_del:
                        del st.session_state[k]
                    if f"submitted_test_{c_num}" in st.session_state:
                        st.session_state[f"submitted_test_{c_num}"] = False
                    if f"test_user_answers_{c_num}" in st.session_state:
                        del st.session_state[f"test_user_answers_{c_num}"]
                    if c_num in st.session_state["completed_chunks"]:
                        st.session_state["completed_chunks"].remove(c_num)
                    if c_num in st.session_state["passed_tests"]:
                        st.session_state["passed_tests"].remove(c_num)
                    save_current_progress()
                    st.success(f"Đã reset Phần {c_num + 1} thành công!")
                    st.rerun()

                elif sub_mode == "⚠️ Ôn các câu sai":
                    st.markdown(f"### Ôn Lại Các Câu Sai Trong Phần {c_num + 1}")
                    chunk_questions_texts = set(q["question"] for q in current_chunk_questions)
                    wrong_in_chunk = [q for q in st.session_state["wrong_questions"] if q.get("question") in chunk_questions_texts]
                    
                    if not wrong_in_chunk:
                        st.info("🎉 Phần này bạn chưa trả lời sai câu nào cả!")
                    else:
                        if f"wrong_chunk_idx_{c_num}" not in st.session_state:
                            st.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            
                        wc_idx = st.session_state[f"wrong_chunk_idx_{c_num}"]
                        if wc_idx >= len(wrong_in_chunk):
                            wc_idx = 0
                            st.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            
                        wc_item = wrong_in_chunk[wc_idx]
                        
                        st.markdown(f"""
                            <div class="main-header-card">
                                <div style="display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 1.1rem; font-weight: 700; color: #f43f5e;">⚠️ Ôn câu sai Phần {c_num + 1}</span>
                                    <span class="badge-topic">Câu sai {wc_idx + 1} / {len(wrong_in_chunk)}</span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown('<div class="question-card">', unsafe_allow_html=True)
                        st.markdown(f"""
                            <div class="question-title">
                                <b>Câu hỏi:</b> {wc_item['question']}
                            </div>
                        """, unsafe_allow_html=True)
                        
                        shuff_data = get_shuffled_options(wc_item, f"shuff_wc_{c_num}_{wc_idx}")
                        options = shuff_data["options"]
                        correct_letter = shuff_data["correct"]
                        
                        wc_storage_key = f"user_ans_wc_{c_num}_{wc_idx}"
                        wc_default_idx = None
                        if st.session_state.get(wc_storage_key) in options:
                            wc_default_idx = options.index(st.session_state.get(wc_storage_key))
                            
                        wc_choice = st.radio("Chọn đáp án của bạn:", options, index=wc_default_idx, key=f"radio_wc_{c_num}_{wc_idx}")
                        if wc_choice is not None:
                            st.session_state[wc_storage_key] = wc_choice
                            st.markdown("<br>", unsafe_allow_html=True)
                            if wc_choice.strip().upper().startswith(correct_letter):
                                st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                                st.session_state["wrong_questions"] = [
                                    w for w in st.session_state["wrong_questions"] 
                                    if w.get("question") != wc_item.get("question")
                                ]
                                save_current_progress()
                            else:
                                st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
                        st.markdown('</div>', unsafe_allow_html=True)
                        
                        if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True, key=f"next_wc_{c_num}_{wc_idx}"):
                            if st.session_state[f"wrong_chunk_idx_{c_num}"] < len(wrong_in_chunk) - 1:
                                st.session_state[f"wrong_chunk_idx_{c_num}"] += 1
                            else:
                                st.session_state[f"wrong_chunk_idx_{c_num}"] = 0
                            st.rerun()

        elif mode == "📝 Thi thử (Mock Test)":
            st.title("📝 Chế Độ Thi Thử (Mock Test)")
            st.markdown("---")
            
            all_questions = []
            for sh, ql in sheets_data.items():
                all_questions.extend(ql)
                
            if "mock_started" not in st.session_state:
                st.session_state["mock_started"] = False
                
            if not st.session_state["mock_started"]:
                st.markdown("""
                    <div class="welcome-card" style="text-align: left; padding: 30px;">
                        <h3>📋 Thông tin bài thi thử mô phỏng:</h3>
                        <p style="color: #cbd5e1; font-size: 1.1rem; margin-top: 10px;">- Bài thi gồm <b>50 câu hỏi ngẫu nhiên</b> được trộn đều từ toàn bộ ngân hàng câu hỏi an toàn điện.</p>
                        <p style="color: #cbd5e1; font-size: 1.1rem;">- Đánh giá chính xác năng lực và mức độ sẵn sàng trước kỳ thi chính thức.</p>
                    </div>
                """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                
                col_m1, col_m2, col_m3 = st.columns([1, 1.5, 1])
                with col_m2:
                    if st.button("🚀 Bắt đầu làm bài thi ngay", type="primary", use_container_width=True):
                        st.session_state["mock_started"] = True
                        st.session_state["mock_questions"] = random.sample(all_questions, min(50, len(all_questions)))
                        st.session_state["mock_answers"] = {}
                        keys_to_del = [k for k in st.session_state.keys() if k.startswith("shuff_mock_") or k.startswith("mock_q_")]
                        for k in keys_to_del:
                            del st.session_state[k]
                        st.rerun()
            else:
                mock_qs = st.session_state["mock_questions"]
                st.write(f"Đang làm bài thi thử chính thức ({len(mock_qs)} câu).")
                st.markdown("---")
                
                for i, q in enumerate(mock_qs):
                    st.markdown('<div class="question-card">', unsafe_allow_html=True)
                    st.markdown(f"""
                        <div style="display: flex; justify-content: space-between; margin-bottom: 10px;">
                            <span class="badge-topic">{q['sheet']}</span>
                            <span style="color: #94a3b8; font-weight: 600;">Câu {i+1} / {len(mock_qs)}</span>
                        </div>
                        <div class="question-title" style="font-size: 1.1rem !important;">
                            {q['question']}
                        </div>
                    """, unsafe_allow_html=True)
                    
                    shuff_data = get_shuffled_options(q, f"shuff_mock_{i}")
                    options = shuff_data["options"]
                    
                    mock_ans_key = f"mock_q_{i}"
                    ans = st.radio("Chọn đáp án:", options, index=(options.index(st.session_state[mock_ans_key]) if st.session_state.get(mock_ans_key) in options else None), key=mock_ans_key, label_visibility="collapsed")
                    st.session_state["mock_answers"][i] = ans
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                if st.button("📤 Nộp bài thi thử", type="primary", use_container_width=True):
                    st.success("Đã nộp bài thành công!")
                    st.balloons()
                    if st.button("Làm bài thi mới", use_container_width=True):
                        st.session_state["mock_started"] = False
                        st.rerun()
