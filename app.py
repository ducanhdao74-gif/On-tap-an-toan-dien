import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import json
import os
import random
import time

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Hệ Thống Ôn Tập Trắc Nghiệm - Điện Lực Điện Biên",
    page_icon="⚡",
    layout="wide"
)

# --- KẾT NỐI GOOGLE SHEETS ---
@st.cache_resource
def init_connection():
    creds_dict = dict(st.secrets["gcp_service_account"])
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    return client

# Khởi tạo client kết nối
try:
    client = init_connection()
    sheet = client.open("quiz_progress").sheet1
except Exception as e:
    st.error(f"Lỗi kết nối Google Sheets: Khởi tạo thất bại. Chi tiết: {e}")
    sheet = None

# --- HÀM ĐỌC/LƯU TIẾN ĐỘ QUA GOOGLE SHEETS ---
def load_saved_progress():
    if sheet is None:
        return {}
    try:
        data = sheet.get_all_records()
        res = {}
        for row in data:
            k = row.get("Key")
            v = row.get("Value")
            if k:
                try:
                    res[k] = json.loads(v) if isinstance(v, str) else v
                except:
                    res[k] = v
        return res
    except Exception as e:
        return {}

def save_current_progress():
    if sheet is None:
        return
    try:
        sheet.clear()
        state_data = {
            "wrong_questions": json.dumps(st.session_state.get("wrong_questions", []), ensure_ascii=False),
            "bookmarked_questions": json.dumps(st.session_state.get("bookmarked_questions", []), ensure_ascii=False),
            "spaced_repetition_data": json.dumps(st.session_state.get("spaced_repetition_data", {}), ensure_ascii=False),
        }
        df = pd.DataFrame(list(state_data.items()), columns=["Key", "Value"])
        sheet.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e:
        st.error(f"Lỗi khi lưu dữ liệu lên Google Sheets: {e}")

# --- KHỞI TẠO SESSION STATE TỪ GOOGLE SHEETS ---
if "initialized_from_sheet" not in st.session_state:
    saved_data = load_saved_progress()
    st.session_state["wrong_questions"] = saved_data.get("wrong_questions", [])
    st.session_state["bookmarked_questions"] = saved_data.get("bookmarked_questions", [])
    st.session_state["spaced_repetition_data"] = saved_data.get("spaced_repetition_data", {})
    st.session_state["initialized_from_sheet"] = True

# --- CSS TÙY CHỈNH (CYBERPUNK THEME) ---
st.markdown("""
<style>
    .main-header-card {
        background: linear-gradient(135deg, #1e1b4b 0%, #311042 100%);
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #4c1d95;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
        margin-bottom: 20px;
    }
    .question-card {
        background-color: #0f172a;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #334155;
        margin-bottom: 15px;
    }
    .question-title {
        font-size: 1.15rem;
        font-weight: 600;
        color: #f8fafc;
        margin-bottom: 15px;
    }
    .badge-topic {
        background-color: #4f46e5;
        color: white;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 500;
    }
</style>
""", unsafe_allow_html=True)

def scroll_to_top():
    st.markdown("<script>window.scrollTo({top: 0, behavior: 'smooth'});</script>", unsafe_allow_html=True)

def get_shuffled_options(q_item, key_suffix):
    shuff_key = f"shuff_{key_suffix}"
    if shuff_key not in st.session_state:
        opts = q_item["options"].copy()
        random.shuffle(opts)
        correct_opt_text = next((o for o in q_item["options"] if o.strip().upper().startswith(q_item["correct"])), "")
        st.session_state[shuff_key] = {
            "options": opts,
            "correct": q_item["correct"],
            "correct_text": correct_opt_text
        }
    return st.session_state[shuff_key]

def update_spaced_repetition(question_text, is_correct):
    sr_data = st.session_state.get("spaced_repetition_data", {})
    if question_text not in sr_data:
        sr_data[question_text] = {"box": 1, "next_review": time.time()}
    
    item = sr_data[question_text]
    if is_correct:
        item["box"] = min(item["box"] + 1, 5)
        days_interval = [0, 1, 3, 7, 14, 30][item["box"]]
        item["next_review"] = time.time() + (days_interval * 86400)
    else:
        item["box"] = 1
        item["next_review"] = time.time()
        
    st.session_state["spaced_repetition_data"] = sr_data
    save_current_progress()

# --- TẢI DỮ LIỆU EXCEL (HỖ TRỢ KHO HƠN 1000 CÂU HỎI) ---
st.sidebar.title("⚡ Menu Điều Hướng")
uploaded_file = st.sidebar.file_uploader("📂 Tải lên tệp Excel câu hỏi (.xlsx)", type=["xlsx"])

sheets_data = {}
if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name)
            q_list = []
            for _, row in df.iterrows():
                # Giả định cấu trúc cột: Câu hỏi, Đáp án A, B, C, D, Đáp án đúng
                q_text = str(row.iloc[0])
                opts = [str(row.iloc[1]), str(row.iloc[2]), str(row.iloc[3]), str(row.iloc[4])]
                corr = str(row.iloc[5]).strip().upper()
                if q_text and q_text != "nan":
                    q_list.append({
                        "question": q_text,
                        "options": [f"A. {opts[0]}", f"B. {opts[1]}", f"C. {opts[2]}", f"D. {opts[3]}"],
                        "correct": corr[0] if corr else "A"
                    })
            if q_list:
                sheets_data[sheet_name] = q_list
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file Excel: {e}")

# Dữ liệu dự phòng nếu chưa tải file
if not sheets_data:
    sheets_data = {
        "Chương Mẫu": [
            {
                "question": "Vui lòng tải tệp Excel chứa câu hỏi lên thanh bên trái để bắt đầu ôn tập toàn bộ hệ thống.",
                "options": ["A. Đã hiểu", "B. Chưa hiểu", "C. Cần hướng dẫn", "D. Bỏ qua"],
                "correct": "A"
            }
        ]
    }

mode = st.sidebar.radio(
    "Chọn chế độ học:",
    [
        "📖 Ôn tập theo chương",
        "🔄 Ôn lại câu trả lời sai",
        "🧠 Spaced Repetition (Ôn thông minh)",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "⭐ Tất cả câu hỏi cần ghi nhớ",
        "📝 Thi thử (Mock Test)",
        "⚙️ Quản lý kho lưu trữ & Dữ liệu"
    ]
)

# --- NỘI DUNG CÁC CHẾ ĐỘ ---
if mode == "📖 Ôn tập theo chương":
    st.title("📖 Ôn Tập Theo Chương")
    st.markdown("---")
    selected_sheet = st.selectbox("Chọn chương học:", list(sheets_data.keys()))
    q_list = sheets_data[selected_sheet]
    
    if f"q_idx_{selected_sheet}" not in st.session_state:
        st.session_state[f"q_idx_{selected_sheet}"] = 0
        
    idx = st.session_state[f"q_idx_{selected_sheet}"]
    item = q_list[idx]
    
    st.markdown(f"""
        <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">📂 {selected_sheet}</span>
                <span class="badge-topic">Câu {idx + 1} / {len(q_list)}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
        <div class="question-card">
            <div class="question-title">{item['question']}</div>
        </div>
    """, unsafe_allow_html=True)
    
    shuff = get_shuffled_options(item, f"{selected_sheet}_{idx}")
    options = shuff["options"]
    correct_letter = shuff["correct"]
    
    ans_key = f"user_ans_{selected_sheet}_{idx}"
    answered_flag = f"answered_{selected_sheet}_{idx}"
    is_answered = st.session_state.get(answered_flag, False)
    
    def_idx = None
    if st.session_state.get(ans_key) in options:
        def_idx = options.index(st.session_state.get(ans_key))
        
    if not is_answered:
        chosen = st.radio("Chọn đáp án:", options, index=def_idx, key=f"radio_{selected_sheet}_{idx}", label_visibility="collapsed")
        if chosen is not None:
            st.session_state[ans_key] = chosen
            st.session_state[answered_flag] = True
            is_correct = chosen.strip().upper().startswith(correct_letter)
            update_spaced_repetition(item["question"], is_correct)
            if not is_correct:
                if not any(w.get("question") == item["question"] for w in st.session_state["wrong_questions"]):
                    st.session_state["wrong_questions"].append(item)
                    save_current_progress()
            st.rerun()
    else:
        saved_choice = st.session_state.get(ans_key)
        st.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_choice}</b></p>", unsafe_allow_html=True)
        is_correct = saved_choice.strip().upper().startswith(correct_letter)
        st.markdown("<br>", unsafe_allow_html=True)
        if is_correct:
            st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
        else:
            correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
            st.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
            
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("⬅️ Câu trước", use_container_width=True):
            if st.session_state[f"q_idx_{selected_sheet}"] > 0:
                st.session_state[f"q_idx_{selected_sheet}"] -= 1
            else:
                st.session_state[f"q_idx_{selected_sheet}"] = len(q_list) - 1
            scroll_to_top()
            st.rerun()
    with col2:
        if st.button("Câu tiếp theo ➡️", type="primary", use_container_width=True):
            if st.session_state[f"q_idx_{selected_sheet}"] < len(q_list) - 1:
                st.session_state[f"q_idx_{selected_sheet}"] += 1
            else:
                st.session_state[f"q_idx_{selected_sheet}"] = 0
            scroll_to_top()
            st.rerun()

elif mode == "🔄 Ôn lại câu trả lời sai":
    st.title("🔄 Ôn Lại Câu Trả Lời Sai")
    st.markdown("---")
    wrong_list = st.session_state.get("wrong_questions", [])
    if not wrong_list:
        st.info("Tuyệt vời! Bạn chưa có câu trả lời sai nào cần ôn lại.")
    else:
        st.warning(f"Bạn đang có **{len(wrong_list)}** câu hỏi cần khắc phục.")
        if "wrong_idx" not in st.session_state:
            st.session_state["wrong_idx"] = 0
            
        w_idx = st.session_state["wrong_idx"]
        if w_idx >= len(wrong_list):
            w_idx = 0
            st.session_state["wrong_idx"] = 0
            
        w_item = wrong_list[w_idx]
        st.markdown(f"""
            <div class="question-card">
                <div class="question-title">{w_item['question']}</div>
            </div>
        """, unsafe_allow_html=True)
        
        w_shuff = get_shuffled_options(w_item, f"wrong_{w_idx}")
        options = w_shuff["options"]
        correct_letter = w_shuff["correct"]
        
        w_storage_key = f"user_ans_wrong_{w_idx}"
        w_answered_flag = f"answered_wrong_{w_idx}"
        is_w_answered = st.session_state.get(w_answered_flag, False)
        
        w_def_idx = None
        if st.session_state.get(w_storage_key) in options:
            w_def_idx = options.index(st.session_state.get(w_storage_key))
            
        if not is_w_answered:
            w_choice = st.radio("Chọn đáp án:", options, index=w_def_idx, key=f"radio_wrong_{w_idx}", label_visibility="collapsed")
            if w_choice is not None:
                st.session_state[w_storage_key] = w_choice
                st.session_state[w_answered_flag] = True
                is_correct = w_choice.strip().upper().startswith(correct_letter)
                if is_correct:
                    st.toast("Đã trả lời đúng câu hỏi này!", icon="🎉")
                st.rerun()
        else:
            saved_choice = st.session_state.get(w_storage_key)
            st.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_choice}</b></p>", unsafe_allow_html=True)
            is_correct = saved_choice.strip().upper().startswith(correct_letter)
            st.markdown("<br>", unsafe_allow_html=True)
            if is_correct:
                st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
                if st.button("✨ Xóa khỏi danh sách câu sai (Đã thuộc)", key=f"remove_wrong_{w_idx}", type="primary"):
                    st.session_state["wrong_questions"] = [w for w in st.session_state["wrong_questions"] if w.get("question") != w_item["question"]]
                    keys_to_clean = [k for k in st.session_state.keys() if f"wrong_{w_idx}" in k]
                    for k in keys_to_clean:
                        del st.session_state[k]
                    save_current_progress()
                    st.toast("Đã loại bỏ câu hỏi khỏi danh sách ôn lại!", icon="🗑️")
                    st.rerun()
            else:
                correct_text = next((opt for opt in options if opt.strip().upper().startswith(correct_letter)), "")
                st.error(f"❌ Sai rồi! Đáp án đúng là **{correct_text}**.")
            
        st.markdown("<br>", unsafe_allow_html=True)
        col_w_prev, col_w_next = st.columns([1, 1])
        with col_w_prev:
            if st.button("⬅️ Câu trước", key="w_prev", use_container_width=True):
                if st.session_state["wrong_idx"] > 0:
                    st.session_state["wrong_idx"] -= 1
                else:
                    st.session_state["wrong_idx"] = len(wrong_list) - 1
                scroll_to_top()
                st.rerun()
        with col_w_next:
            if st.button("Câu tiếp theo ➡️", key="w_next", type="primary", use_container_width=True):
                if st.session_state["wrong_idx"] < len(wrong_list) - 1:
                    st.session_state["wrong_idx"] += 1
                else:
                    st.session_state["wrong_idx"] = 0
                scroll_to_top()
                st.rerun()

elif mode == "🧠 Spaced Repetition (Ôn thông minh)":
    st.title("🧠 Spaced Repetition - Ôn Tập Thông Minh")
    st.markdown("---")
    st.info("Hệ thống tự động sắp xếp các câu hỏi dựa trên độ khó và lịch sử trả lời của bạn.")
    
    sr_dict = st.session_state["spaced_repetition_data"]
    if not sr_dict:
        st.warning("Chưa có dữ liệu ôn tập thông minh. Hãy làm các bài trắc nghiệm thông thường trước!")
    else:
        now_time = time.time()
        due_questions = []
        for sheet_name, q_list in sheets_data.items():
            for q in q_list:
                q_txt = q["question"]
                if q_txt in sr_dict:
                    item = sr_dict[q_txt]
                    if item.get("next_review", 0) <= now_time:
                        due_questions.append((q, item))
                else:
                    due_questions.append((q, {"box": 1, "next_review": 0}))
        
        st.write(f"Số câu cần ôn tập trong khung thời gian hiện tại: **{len(due_questions)}** câu.")
        if due_questions:
            if "sr_idx" not in st.session_state:
                st.session_state["sr_idx"] = 0
            
            sr_i = st.session_state["sr_idx"]
            if sr_i >= len(due_questions):
                sr_i = 0
                st.session_state["sr_idx"] = 0
                
            current_q_item, current_sr_info = due_questions[sr_i]
            
            st.markdown(f"""
                <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <span style="font-size: 1.05rem; font-weight: 700; color: #c084fc;">🧠 Hộp trí nhớ (Box {current_sr_info.get('box', 1)})</span>
                        <span class="badge-topic">Thẻ {sr_i + 1} / {len(due_questions)}</span>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            st.markdown(f"""
                <div class="question-card">
                    <div class="question-title">{current_q_item['question']}</div>
                </div>
            """, unsafe_allow_html=True)
            
            sr_shuff = get_shuffled_options(current_q_item, f"shuff_sr_{sr_i}")
            sr_opts = sr_shuff["options"]
            sr_corr = sr_shuff["correct"]
            
            sr_ans_key = f"user_ans_sr_{sr_i}"
            sr_answered_key = f"answered_sr_{sr_i}"
            is_sr_done = st.session_state.get(sr_answered_key, False)
            
            sr_def_idx = None
            if st.session_state.get(sr_ans_key) in sr_opts:
                sr_def_idx = sr_opts.index(st.session_state.get(sr_ans_key))
                
            if not is_sr_done:
                chosen_sr = st.radio("Chọn đáp án:", sr_opts, index=sr_def_idx, key=f"radio_sr_{sr_i}", label_visibility="collapsed")
                if chosen_sr is not None:
                    st.session_state[sr_ans_key] = chosen_sr
                    st.session_state[sr_answered_key] = True
                    is_corr = chosen_sr.strip().upper().startswith(sr_corr)
                    update_spaced_repetition(current_q_item["question"], is_corr)
                    st.rerun()
            else:
                saved_sr_choice = st.session_state.get(sr_ans_key)
                st.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_sr_choice}</b></p>", unsafe_allow_html=True)
                is_corr = saved_sr_choice.strip().upper().startswith(sr_corr)
                st.markdown("<br>", unsafe_allow_html=True)
                if is_corr:
                    st.success("🎉 Chính xác! Thẻ đã được đẩy lên cấp độ cao hơn.")
                else:
                    corr_txt = next((o for o in sr_opts if o.strip().upper().startswith(sr_corr)), "")
                    st.error(f"❌ Sai rồi! Đáp án đúng là **{corr_txt}**.")
                    
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Thẻ tiếp theo ➡️", key="next_sr_card", type="primary", use_container_width=True):
                if st.session_state["sr_idx"] < len(due_questions) - 1:
                    st.session_state["sr_idx"] += 1
                else:
                    st.session_state["sr_idx"] = 0
                scroll_to_top()
                st.rerun()

elif mode == "📂 Ôn gộp tất cả (50 câu/phần)":
    st.title("📂 Ôn Tập Gộp Toàn Bộ Hệ Thống")
    st.markdown("---")
    
    all_flat_questions = []
    for sh, q_list in sheets_data.items():
        all_flat_questions.extend(q_list)
        
    chunk_size = 50
    total_chunks = (len(all_flat_questions) + chunk_size - 1) // chunk_size
    
    chunk_options = [f"Phần {i+1}: Câu {i*chunk_size + 1} đến {min((i+1)*chunk_size, len(all_flat_questions))}" for i in range(total_chunks)]
    selected_chunk_label = st.selectbox("Chọn phần ôn tập:", chunk_options)
    chunk_idx = chunk_options.index(selected_chunk_label)
    
    start_idx = chunk_idx * chunk_size
    end_idx = min((chunk_idx + 1) * chunk_size, len(all_flat_questions))
    current_chunk_questions = all_flat_questions[start_idx:end_idx]
    
    if f"chunk_q_idx_{chunk_idx}" not in st.session_state:
        st.session_state[f"chunk_q_idx_{chunk_idx}"] = 0
        
    c_idx = st.session_state[f"chunk_q_idx_{chunk_idx}"]
    if c_idx >= len(current_chunk_questions):
        c_idx = 0
        st.session_state[f"chunk_q_idx_{chunk_idx}"] = 0
        
    cq_item = current_chunk_questions[c_idx]
    
    st.markdown(f"""
        <div class="main-header-card" style="margin-top: 0; margin-bottom: 20px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <span style="font-size: 1.05rem; font-weight: 700; color: #38bdf8;">📂 {selected_chunk_label}</span>
                <span class="badge-topic">Câu {c_idx + 1} / {len(current_chunk_questions)}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
        <div class="question-card">
            <div class="question-title">{cq_item['question']}</div>
        </div>
    """, unsafe_allow_html=True)
    
    c_shuff = get_shuffled_options(cq_item, f"shuff_chunk_{chunk_idx}_{c_idx}")
    c_opts = c_shuff["options"]
    c_corr = c_shuff["correct"]
    
    c_ans_key = f"user_ans_chunk_{chunk_idx}_{c_idx}"
    c_ans_flag = f"answered_chunk_{chunk_idx}_{c_idx}"
    is_c_answered = st.session_state.get(c_ans_flag, False)
    
    c_def_idx = None
    if st.session_state.get(c_ans_key) in c_opts:
        c_def_idx = c_opts.index(st.session_state.get(c_ans_key))
        
    if not is_c_answered:
        chosen_c = st.radio("Đáp án:", c_opts, index=c_def_idx, key=f"radio_chunk_{chunk_idx}_{c_idx}", label_visibility="collapsed")
        if chosen_c is not None:
            st.session_state[c_ans_key] = chosen_c
            st.session_state[c_ans_flag] = True
            is_correct = chosen_c.strip().upper().startswith(c_corr)
            update_spaced_repetition(cq_item["question"], is_correct)
            if not is_correct:
                if not any(w.get("question") == cq_item["question"] for w in st.session_state["wrong_questions"]):
                    st.session_state["wrong_questions"].append(cq_item)
                    save_current_progress()
            st.rerun()
    else:
        saved_c_choice = st.session_state.get(c_ans_key)
        st.markdown(f"<p style='color: #cbd5e1; font-size: 1.05rem;'>Đã chọn: <b>{saved_c_choice}</b></p>", unsafe_allow_html=True)
        is_correct = saved_c_choice.strip().upper().startswith(c_corr)
        st.markdown("<br>", unsafe_allow_html=True)
        if is_correct:
            st.success(f"🎉 Chính xác! Đáp án đúng là {c_corr}.")
        else:
            corr_text = next((o for o in c_opts if o.strip().upper().startswith(c_corr)), "")
            st.error(f"❌ Sai rồi! Đáp án đúng là **{corr_text}**.")
            
    st.markdown("<br>", unsafe_allow_html=True)
    col_cp1, col_cp2 = st.columns([1, 1])
    with col_cp1:
        if st.button("⬅️ Câu trước", key="chunk_prev", use_container_width=True):
            if st.session_state[f"chunk_q_idx_{chunk_idx}"] > 0:
                st.session_state[f"chunk_q_idx_{chunk_idx}"] -= 1
            else:
                st.session_state[f"chunk_q_idx_{chunk_idx}"] = len(current_chunk_questions) - 1
            scroll_to_top()
            st.rerun()
    with col_cp2:
        if st.button("Câu tiếp theo ➡️", key="chunk_next", type="primary", use_container_width=True):
            if st.session_state[f"chunk_q_idx_{chunk_idx}"] < len(current_chunk_questions) - 1:
                st.session_state[f"chunk_q_idx_{chunk_idx}"] += 1
            else:
                st.session_state[f"chunk_q_idx_{chunk_idx}"] = 0
            scroll_to_top()
            st.rerun()

elif mode == "⭐ Tất cả câu hỏi cần ghi nhớ":
    st.title("⭐ Danh Sách Câu Hỏi Đã Đánh Dấu Ghi Nhớ")
    st.markdown("---")
    bookmarks = st.session_state.get("bookmarked_questions", [])
    if not bookmarks:
        st.info("Bạn chưa đánh dấu câu hỏi nào.")
    else:
        st.write(f"Đang lưu trữ **{len(bookmarks)}** câu hỏi quan trọng.")
        for idx, bm_item in enumerate(bookmarks):
            with st.expander(f"Câu {idx + 1}: {bm_item.get('question', '')[:80]}..."):
                st.markdown(f"**Câu hỏi đầy đủ:** {bm_item.get('question')}")
                st.markdown("**Các lựa chọn:**")
                for opt in bm_item.get("options", []):
                    st.markdown(f"- {opt}")
                st.markdown(f"**Đáp án đúng chuẩn:** {bm_item.get('correct')}")
                if st.button("🗑️ Xóa khỏi danh sách ghi nhớ", key=f"remove_bm_{idx}"):
                    st.session_state["bookmarked_questions"] = [b for b in st.session_state["bookmarked_questions"] if b.get("question") != bm_item.get("question")]
                    save_current_progress()
                    st.success("Đã xóa khỏi danh sách!")
                    st.rerun()

elif mode == "📝 Thi thử (Mock Test)":
    st.title("📝 Chế Độ Thi Thử (Mock Test)")
    st.markdown("---")
    if "mock_test_started" not in st.session_state:
        st.session_state["mock_test_started"] = False
        
    if not st.session_state["mock_test_started"]:
        test_size = st.slider("Chọn số lượng câu hỏi cho đề thi:", min_value=10, max_value=100, value=30, step=10)
        if st.button("🚀 Bắt đầu làm bài thi", type="primary"):
            all_q_pool = []
            for sh, q_list in sheets_data.items():
                all_q_pool.extend(q_list)
            random.shuffle(all_q_pool)
            st.session_state["mock_questions"] = all_q_pool[:test_size]
            st.session_state["mock_answers"] = {}
            st.session_state["mock_test_started"] = True
            st.rerun()
    else:
        mock_qs = st.session_state["mock_questions"]
        submitted = st.session_state.get("mock_submitted", False)
        
        if not submitted:
            st.markdown("### Bài thi đang diễn ra...")
            for i, mq in enumerate(mock_qs):
                st.markdown(f"**Câu {i+1}:** {mq['question']}")
                m_shuff = get_shuffled_options(mq, f"shuff_mock_{i}")
                m_choice = st.radio("Chọn đáp án:", m_shuff["options"], key=f"radio_mock_{i}", index=None)
                if m_choice:
                    st.session_state["mock_answers"][i] = m_choice
                st.markdown("---")
                
            if st.button("📥 Nộp bài thi", type="primary"):
                st.session_state["mock_submitted"] = True
                st.rerun()
        else:
            st.markdown("### 📊 Kết Quả Bài Thi Thử")
            score = 0
            total_m = len(mock_qs)
            for i, mq in enumerate(mock_qs):
                user_ans = st.session_state["mock_answers"].get(i, "")
                m_shuff = st.session_state.get(f"shuff_mock_{i}", {})
                corr_let = m_shuff.get("correct", "A")
                is_corr = user_ans.strip().upper().startswith(corr_let)
                if is_corr:
                    score += 1
                    
            st.metric("Điểm số của bạn", f"{score} / {total_m}", delta=f"{(score/total_m)*100:.1f}%")
            if st.button("🔄 Làm bài thi mới"):
                st.session_state["mock_test_started"] = False
                st.session_state["mock_submitted"] = False
                st.session_state["mock_answers"] = {}
                st.rerun()

elif mode == "⚙️ Quản lý kho lưu trữ & Dữ liệu":
    st.title("⚙️ Quản Lý Kho Lưu Trữ & Dữ Liệu (Google Sheets)")
    st.markdown("---")
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.markdown("### 📥 Tải dữ liệu từ Sheet")
        if st.button("Làm mới & đồng bộ dữ liệu từ Google Sheets", use_container_width=True):
            saved_data = load_saved_progress()
            st.session_state["wrong_questions"] = saved_data.get("wrong_questions", [])
            st.session_state["bookmarked_questions"] = saved_data.get("bookmarked_questions", [])
            st.session_state["spaced_repetition_data"] = saved_data.get("spaced_repetition_data", {})
            st.success("Đã đồng bộ dữ liệu thành công từ Google Sheets!")
            st.rerun()
            
    with col_m2:
        st.markdown("### 🗑️ Đặt lại hệ thống")
        if st.button("⚠️ Xóa toàn bộ tiến độ trên Google Sheets", type="primary", use_container_width=True):
            if sheet is not None:
                sheet.clear()
            for key in list(st.session_state.keys()):
                if key != "initialized_from_sheet":
                    del st.session_state[key]
            st.success("Đã xóa và đặt lại toàn bộ tiến độ thành công!")
            st.rerun()
