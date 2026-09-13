import subprocess
import threading
import time
import os

subprocess.run(["pkill", "-f", "streamlit"])

app_code = '''
import streamlit as st
import openpyxl
import re
import os
import random

st.set_page_config(page_title="Hệ Thống Ôn Thi & Thi Thử An Toàn Điện", layout="wide")

# Tùy chỉnh CSS để tăng cỡ chữ đáp án lên 18px
st.markdown("""
<style>
    /* Tăng cỡ chữ cho các nhãn (label) của radio button (đáp án A, B, C, D) */
    div.stRadio > label {
        font-size: 18px !important;
        font-weight: 500 !important;
    }
    div.stRadio label p {
        font-size: 18px !important;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data_with_color(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    data_bank = {}
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        rows = []
        correct_flags = []
        
        for row in ws.iter_rows(min_col=1, max_col=1):
            cell = row[0]
            val = cell.value
            if val is not None:
                val_str = str(val).strip()
                rows.append(val_str)
                
                color = cell.font.color if cell.font else None
                is_red = False
                if color and color.rgb:
                    rgb_str = str(color.rgb).upper()
                    if "FF0000" in rgb_str or rgb_str.endswith("FF0000") or "RED" in str(color.theme):
                        is_red = True
                correct_flags.append(is_red)
                
        questions = []
        current_q = None
        current_opts = []
        current_correct_opt = None
        
        for idx, row_str in enumerate(rows):
            is_red = correct_flags[idx]
            
            if re.match(r'^Câu\\s*\\d+[:\\.]', row_str, re.IGNORECASE):
                if current_q and current_opts:
                    questions.append({
                        "id": len(questions),
                        "sheet": sheet,
                        "question": current_q, 
                        "options": current_opts, 
                        "correct": current_correct_opt
                    })
                current_q = row_str
                current_opts = []
                current_correct_opt = None
            elif re.match(r'^[A-D]\\.', row_str):
                current_opts.append(row_str)
                if is_red:
                    current_correct_opt = row_str
            else:
                if current_q and not current_opts:
                    current_q += " " + row_str
                elif current_opts:
                    current_opts[-1] += " " + row_str
                    if is_red and current_correct_opt is None:
                        current_correct_opt = current_opts[-1]
                        
        if current_q and current_opts:
            questions.append({
                "id": len(questions),
                "sheet": sheet,
                "question": current_q, 
                "options": current_opts, 
                "correct": current_correct_opt
            })
            
        data_bank[sheet] = questions
    return data_bank

file_path = 'PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx'
if os.path.exists(file_path):
    bank = load_data_with_color(file_path)
else:
    st.error("Không tìm thấy file Excel!")
    st.stop()

if 'learned' not in st.session_state:
    st.session_state.learned = {sheet: set() for sheet in bank.keys()}
if 'wrong_questions' not in st.session_state:
    st.session_state.wrong_questions = {sheet: [] for sheet in bank.keys()}
if 'history_answers' not in st.session_state:
    st.session_state.history_answers = {sheet: {} for sheet in bank.keys()}
if 'current_view_idx' not in st.session_state:
    st.session_state.current_view_idx = {sheet: 0 for sheet in bank.keys()}

if 'shuffled_combined_questions' not in st.session_state:
    all_combined = []
    for sheet, q_list in bank.items():
        all_combined.extend(q_list)
    random.seed(42)
    random.shuffle(all_combined)
    for new_id, q in enumerate(all_combined):
        q['id'] = new_id
    st.session_state.shuffled_combined_questions = all_combined

all_combined_questions = st.session_state.shuffled_combined_questions

if 'combined_learned' not in st.session_state:
    st.session_state.combined_learned = set()
if 'combined_wrong' not in st.session_state:
    st.session_state.combined_wrong = []
if 'combined_history' not in st.session_state:
    st.session_state.combined_history = {}
if 'combined_view_idx' not in st.session_state:
    st.session_state.combined_view_idx = 0
if 'completed_parts' not in st.session_state:
    st.session_state.completed_parts = set()

# Khởi tạo session state lưu index câu đang xem trong chế độ "Làm lại câu sai của phần"
if 'combined_wrong_view_idx' not in st.session_state:
    st.session_state.combined_wrong_view_idx = {}

st.sidebar.title("⚡ Menu Ôn Tập")
mode = st.sidebar.radio("Chọn chế độ:", [
    "📖 Ôn tập theo chuyên đề", 
    "🔁 Ôn lại câu trả lời sai", 
    "📚 Ôn gộp tất cả (50 câu/phần)",
    "📝 Thi thử (Mock Test)"
])

sheet_names = list(bank.keys())

if mode == "📖 Ôn tập theo chuyên đề":
    selected_sheet = st.sidebar.selectbox("📂 Chọn Chuyên Đề", sheet_names)
    all_questions = bank[selected_sheet]
    st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
    
    learned_indices = st.session_state.learned[selected_sheet]
    
    st.sidebar.markdown("---")
    st.sidebar.metric("Tổng số câu", len(all_questions))
    st.sidebar.metric("Đã làm", len(learned_indices))
    
    if st.sidebar.button("🔄 Đặt lại tiến độ chuyên đề này"):
        st.session_state.learned[selected_sheet] = set()
        st.session_state.wrong_questions[selected_sheet] = []
        st.session_state.history_answers[selected_sheet] = {}
        st.session_state.current_view_idx[selected_sheet] = 0
        st.rerun()

    cur_idx = st.session_state.current_view_idx[selected_sheet]
    if cur_idx >= len(all_questions):
        cur_idx = len(all_questions) - 1
        st.session_state.current_view_idx[selected_sheet] = cur_idx

    history_dict = st.session_state.history_answers[selected_sheet]

    q_data = all_questions[cur_idx]
    
    st.subheader(f"Chuyên đề: {selected_sheet} (Câu {cur_idx + 1}/{len(all_questions)})")
    st.progress((cur_idx + 1) / len(all_questions))
    st.markdown("---")
    
    st.markdown(f"### {q_data['question']}")
    
    def on_choose_answer():
        chosen = st.session_state[f"study_{selected_sheet}_{cur_idx}"]
        correct_ans = q_data.get('correct')
        if correct_ans:
            st.session_state.learned[selected_sheet].add(cur_idx)
            is_corr = (chosen.strip() == correct_ans.strip())
            
            st.session_state.history_answers[selected_sheet][cur_idx] = {
                "question": q_data['question'],
                "user_choice": chosen,
                "correct_ans": correct_ans,
                "is_correct": is_corr
            }
            
            if is_corr:
                if cur_idx in st.session_state.wrong_questions[selected_sheet]:
                    st.session_state.wrong_questions[selected_sheet].remove(cur_idx)
            else:
                if cur_idx not in st.session_state.wrong_questions[selected_sheet]:
                    st.session_state.wrong_questions[selected_sheet].append(cur_idx)

    default_val = history_dict[cur_idx]['user_choice'] if cur_idx in history_dict else None
    opt_index = q_data['options'].index(default_val) if default_val in q_data['options'] else None

    st.radio(
        "Chọn đáp án của bạn:", 
        q_data['options'], 
        key=f"study_{selected_sheet}_{cur_idx}", 
        index=opt_index,
        on_change=on_choose_answer
    )

    if cur_idx in history_dict:
        h_info = history_dict[cur_idx]
        st.markdown("---")
        if h_info['is_correct']:
            st.success(f"✔️ Bạn đã chọn đúng: **{h_info['user_choice']}**")
        else:
            st.error(f"❌ Bạn chọn: `{h_info['user_choice']}` | Đáp án chuẩn: **{h_info['correct_ans']}**")

    st.markdown("---")
    nav_col1, nav_col2 = st.columns(2)
    with nav_col1:
        if cur_idx > 0:
            if st.button("⬅️ Quay lại câu trước"):
                st.session_state.current_view_idx[selected_sheet] -= 1
                st.rerun()
    with nav_col2:
        if cur_idx < len(all_questions) - 1:
            if st.button("Bỏ qua / Sang câu tiếp theo ➡️"):
                st.session_state.current_view_idx[selected_sheet] += 1
                st.rerun()

elif mode == "🔁 Ôn lại câu trả lời sai":
    selected_sheet = st.sidebar.selectbox("📂 Chọn Chuyên Đề", sheet_names)
    all_questions = bank[selected_sheet]
    st.title("🔁 Ôn Lập Lại Các Câu Trả Lời Sai")
    wrong_list = st.session_state.wrong_questions[selected_sheet]
    
    if not wrong_list:
        st.info("🎉 Tuyệt vời! Bạn chưa có câu trả lời sai nào trong chuyên đề này.")
    else:
        st.warning(f"Bạn đang có **{len(wrong_list)} câu** làm sai cần ôn tập lại.")
        for idx in wrong_list:
            q_data = all_questions[idx]
            st.markdown(f"---")
            st.markdown(f"### {q_data['question']}")
            ans = st.radio("Chọn đáp án:", q_data['options'], key=f"retry_{selected_sheet}_{idx}")
            if st.button(f"Kiểm tra câu hỏi này", key=f"btn_retry_{idx}"):
                correct_ans = q_data.get('correct')
                if ans.strip() == correct_ans.strip():
                    st.success(f"🎉 Chính xác! Bạn đã sửa thành công.")
                    st.session_state.wrong_questions[selected_sheet].remove(idx)
                    st.rerun()
                else:
                    st.error(f"❌ Vẫn chưa đúng. Đáp án chuẩn là: **{correct_ans}**")

elif mode == "📚 Ôn gộp tất cả (50 câu/phần)":
    st.title("📚 Ôn Gộp Tất Cả Chuyên Đề (Trộn Đều & Chia Phần 50 Câu)")
    
    chunk_size = 50
    total_q = len(all_combined_questions)
    chunks = [all_combined_questions[i:i + chunk_size] for i in range(0, total_q, chunk_size)]
    
    part_options = []
    for i in range(len(chunks)):
        sheets_in_chunk = list(set(q['sheet'] for q in chunks[i]))
        sheets_str = f"Hỗn hợp {len(sheets_in_chunk)} chuyên đề"
        status = " - ✅ Done" if i in st.session_state.completed_parts else ""
        part_options.append(f"Phần {i+1} ({sheets_str}){status}")
    
    selected_part_str = st.sidebar.selectbox("📂 Chọn Phần Ôn Tập", part_options)
    part_idx = int(selected_part_str.split(" ")[1]) - 1
    current_chunk = chunks[part_idx]

    sub_mode = st.radio("Chọn chế độ cho phần này:", [
        "📖 Ôn tập từng câu trong phần", 
        "📝 Bài kiểm tra chốt kiến thức phần này",
        "🔄 Làm lại phần này (Ôn tập lại từ đầu)",
        "🔁 Làm lại các câu sai trong phần này"
    ], horizontal=True)
    st.markdown("---")
    
    if sub_mode == "📖 Ôn tập từng câu trong phần":
        st.subheader(f"Đang ôn: {selected_part_str}")
        
        st.sidebar.markdown("---")
        st.sidebar.metric("Tổng số câu của phần", len(current_chunk))

        c_idx = st.session_state.combined_view_idx
        if c_idx >= len(current_chunk):
            c_idx = len(current_chunk) - 1
            st.session_state.combined_view_idx = c_idx

        c_history = st.session_state.combined_history

        q_data = current_chunk[c_idx]
        global_id = q_data['id']
        
        st.markdown(f"*(Thuộc chuyên đề: **{q_data['sheet']}**)*")
        st.markdown(f"### Câu {c_idx + 1}/{len(current_chunk)} (Toàn hệ thống #{global_id + 1}): {q_data['question']}")
        
        def on_choose_combined():
            chosen = st.session_state[f"combined_study_{global_id}"]
            correct_ans = q_data.get('correct')
            if correct_ans:
                st.session_state.combined_learned.add(global_id)
                is_corr = (chosen.strip() == correct_ans.strip())
                
                st.session_state.combined_history[global_id] = {
                    "sheet": q_data['sheet'],
                    "question": q_data['question'],
                    "user_choice": chosen,
                    "correct_ans": correct_ans,
                    "is_correct": is_corr
                }
                
                if is_corr:
                    if global_id in st.session_state.combined_wrong:
                        st.session_state.combined_wrong.remove(global_id)
                else:
                    if global_id not in st.session_state.combined_wrong:
                        st.session_state.combined_wrong.append(global_id)

        default_choice = c_history[global_id]['user_choice'] if global_id in c_history else None
        c_opt_index = q_data['options'].index(default_choice) if default_choice in q_data['options'] else None

        st.radio(
            "Chọn đáp án:", 
            q_data['options'], 
            key=f"combined_study_{global_id}", 
            index=c_opt_index,
            on_change=on_choose_combined
        )

        if global_id in c_history:
            h_info = c_history[global_id]
            st.markdown("---")
            if h_info['is_correct']:
                st.success(f"✔️ Bạn đã chọn đúng: **{h_info['user_choice']}**")
            else:
                st.error(f"❌ Bạn chọn: `{h_info['user_choice']}` | Đáp án chuẩn: **{h_info['correct_ans']}**")

        st.markdown("---")
        nav_col1, nav_col2 = st.columns(2)
        with nav_col1:
            if c_idx > 0:
                if st.button("⬅️ Quay lại câu trước", key="prev_comb"):
                    st.session_state.combined_view_idx -= 1
                    st.rerun()
        with nav_col2:
            if c_idx < len(current_chunk) - 1:
                if st.button("Bỏ qua / Sang câu tiếp theo ➡️", key="next_comb"):
                    st.session_state.combined_view_idx += 1
                    st.rerun()
                    
    elif sub_mode == "📝 Bài kiểm tra chốt kiến thức phần này":
        st.subheader(f"📝 Bài Kiểm Tra Chốt Kiến Thức - {selected_part_str}")
        st.info("Bài kiểm tra gồm toàn bộ câu hỏi hỗn hợp trong phần này để đánh giá mức độ nhớ kiến thức của bạn.")
        
        exam_key = f"combined_exam_sub_{part_idx}"
        if exam_key not in st.session_state or st.sidebar.button("🔄 Làm lại bài kiểm tra phần này"):
            st.session_state[exam_key] = {
                "submitted": False,
                "answers": {}
            }
            if part_idx in st.session_state.completed_parts:
                st.session_state.completed_parts.remove(part_idx)
            st.rerun()
            
        exam_state = st.session_state[exam_key]
        
        if not exam_state["submitted"]:
            with st.form(f"form_part_{part_idx}"):
                for i, q in enumerate(current_chunk):
                    st.markdown(f"**Câu {i+1}: {q['question']}** *(Chuyên đề: {q['sheet']})*")
                    choice = st.radio("Lựa chọn:", q['options'], key=f"part_q_{part_idx}_{i}", index=None)
                    exam_state["answers"][i] = choice
                    st.markdown("---")
                
                submitted_exam = st.form_submit_button("🏁 Nộp Bài Kiểm Tra Phần", type="primary")
                if submitted_exam:
                    exam_state["submitted"] = True
                    st.session_state.completed_parts.add(part_idx)
                    st.rerun()
        else:
            st.subheader("📊 Kết Quả Bài Kiểm Tra Phần")
            score = 0
            total = len(current_chunk)
            
            for i, q in enumerate(current_chunk):
                user_ans = exam_state["answers"].get(i)
                correct_ans = q.get('correct')
                
                st.markdown(f"**Câu {i+1}: {q['question']}** *(Chuyên đề: {q['sheet']})*")
                if user_ans and correct_ans and user_ans.strip() == correct_ans.strip():
                    st.success(f"Bạn chọn: {user_ans} (Đúng)")
                    score += 1
                else:
                    st.error(f"Bạn chọn: {user_ans} | Đáp án chuẩn: **{correct_ans}**")
                st.markdown("---")
                
            st.metric(label="Kết quả tổng kết phần", value=f"{score} / {total} câu đúng ({score/total*100:.1f}%)")
            if st.button("🔄 Làm lại bài kiểm tra này"):
                exam_state["submitted"] = False
                if part_idx in st.session_state.completed_parts:
                    st.session_state.completed_parts.remove(part_idx)
                st.rerun()

    elif sub_mode == "🔄 Làm lại phần này (Ôn tập lại từ đầu)":
        st.subheader(f"🔄 Đặt Lại & Ôn Tập Lại Phần Này - {selected_part_str}")
        st.warning("Tính năng này sẽ xóa lịch sử câu trả lời đã làm của riêng phần này để ông ôn tập lại từ đầu như mới.")
        if st.button("🚀 Xác nhận reset và bắt đầu ôn lại phần này", type="primary"):
            for q in current_chunk:
                gid = q['id']
                if gid in st.session_state.combined_history:
                    del st.session_state.combined_history[gid]
                if gid in st.session_state.combined_learned:
                    st.session_state.combined_learned.remove(gid)
                if gid in st.session_state.combined_wrong:
                    st.session_state.combined_wrong.remove(gid)
            st.session_state.combined_view_idx = 0
            if part_idx in st.session_state.completed_parts:
                st.session_state.completed_parts.remove(part_idx)
            st.success("Đã reset thành công! Chuyển về câu đầu tiên...")
            st.rerun()

    else: # "🔁 Làm lại các câu sai trong phần này"
        st.subheader(f"🔁 Làm Lại Các Câu Làm Sai - {selected_part_str}")
        
        # Lọc ra các câu trong chunk hiện tại mà đang nằm trong danh sách combined_wrong
        chunk_wrong_ids = [q['id'] for q in current_chunk if q['id'] in st.session_state.combined_wrong]
        chunk_wrong_questions = [q for q in current_chunk if q['id'] in st.session_state.combined_wrong]
        
        if not chunk_wrong_questions:
            st.info("🎉 Tuyệt vời! Bạn không có câu trả lời sai nào trong phần này.")
        else:
            st.warning(f"Bạn đang có **{len(chunk_wrong_questions)} câu** làm sai trong phần này cần luyện tập lại.")
            
            if part_idx not in st.session_state.combined_wrong_view_idx:
                st.session_state.combined_wrong_view_idx[part_idx] = 0
            
            w_idx = st.session_state.combined_wrong_view_idx[part_idx]
            if w_idx >= len(chunk_wrong_questions):
                w_idx = len(chunk_wrong_questions) - 1
                st.session_state.combined_wrong_view_idx[part_idx] = w_idx
                
            q_data = chunk_wrong_questions[w_idx]
            global_id = q_data['id']
            c_history = st.session_state.combined_history
            
            st.markdown(f"*(Thuộc chuyên đề: **{q_data['sheet']}**)*")
            st.markdown(f"### Câu sai {w_idx + 1}/{len(chunk_wrong_questions)} (Toàn hệ thống #{global_id + 1}): {q_data['question']}")
            
            def on_choose_combined_wrong():
                chosen = st.session_state[f"combined_wrong_study_{global_id}"]
                correct_ans = q_data.get('correct')
                if correct_ans:
                    is_corr = (chosen.strip() == correct_ans.strip())
                    
                    st.session_state.combined_history[global_id] = {
                        "sheet": q_data['sheet'],
                        "question": q_data['question'],
                        "user_choice": chosen,
                        "correct_ans": correct_ans,
                        "is_correct": is_corr
                    }
                    
                    if is_corr:
                        if global_id in st.session_state.combined_wrong:
                            st.session_state.combined_wrong.remove(global_id)
                    else:
                        if global_id not in st.session_state.combined_wrong:
                            st.session_state.combined_wrong.append(global_id)

            default_choice = c_history[global_id]['user_choice'] if global_id in c_history else None
            c_opt_index = q_data['options'].index(default_choice) if default_choice in q_data['options'] else None

            st.radio(
                "Chọn đáp án:", 
                q_data['options'], 
                key=f"combined_wrong_study_{global_id}", 
                index=c_opt_index,
                on_change=on_choose_combined_wrong
            )

            if global_id in c_history:
                h_info = c_history[global_id]
                st.markdown("---")
                if h_info['is_correct']:
                    st.success(f"✔️ Bạn đã chọn đúng: **{h_info['user_choice']}**")
                else:
                    st.error(f"❌ Bạn chọn: `{h_info['user_choice']}` | Đáp án chuẩn: **{h_info['correct_ans']}**")

            st.markdown("---")
            nav_col1, nav_col2 = st.columns(2)
            with nav_col1:
                if w_idx > 0:
                    if st.button("⬅️ Quay lại câu trước", key="prev_comb_wrong"):
                        st.session_state.combined_wrong_view_idx[part_idx] -= 1
                        st.rerun()
            with nav_col2:
                if w_idx < len(chunk_wrong_questions) - 1:
                    if st.button("Bỏ qua / Sang câu tiếp theo ➡️", key="next_comb_wrong"):
                        st.session_state.combined_wrong_view_idx[part_idx] += 1
                        st.rerun()

elif mode == "📝 Thi thử (Mock Test)":
    st.title("📝 Chế Độ Thi Thử Trắc Nghiệm")
    num_questions = st.sidebar.slider("Chọn số lượng câu hỏi đề thi:", 5, 50, 20)
    
    if 'exam_questions' not in st.session_state or st.sidebar.button("🔄 Tạo đề thi mới"):
        st.session_state.exam_questions = random.sample(all_combined_questions, min(num_questions, len(all_combined_questions)))
        st.session_state.submitted = False
        st.session_state.user_answers = {}

    exam_qs = st.session_state.exam_questions
    
    if not st.session_state.get('submitted', False):
        with st.form("exam_form"):
            for i, q in enumerate(exam_qs):
                st.markdown(f"**Câu {i+1}: {q['question']}** *(Chuyên đề: {q['sheet']})*")
                choice = st.radio("Lựa chọn:", q['options'], key=f"exam_{i}", index=None)
                st.session_state.user_answers[i] = choice
                st.markdown("---")
            
            submitted = st.form_submit_button("🏁 Nộp Bài Thi", type="primary")
            if submitted:
                st.session_state.submitted = True
                st.rerun()
    else:
        st.subheader("📊 Kết Quả Bài Thi Thử")
        score = 0
        total = len(exam_qs)
        
        for i, q in enumerate(exam_qs):
            user_ans = st.session_state.user_answers.get(i)
            correct_ans = q.get('correct')
            
            st.markdown(f"**Câu {i+1}: {q['question']}**")
            if user_ans and correct_ans and user_ans.strip() == correct_ans.strip():
                st.success(f"Bạn chọn: {user_ans} (Đúng)")
                score += 1
            else:
                st.error(f"Bạn chọn: {user_ans} | Đáp án chuẩn: **{correct_ans}**")
            st.markdown("---")
            
        st.metric(label="Điểm số tổng kết", value=f"{score} / {total} câu đúng ({score/total*100:.1f}%)")
        if st.button("🔄 Làm đề thi mới"):
            st.session_state.submitted = False
            st.rerun()
'''

with open("app.py", "w", encoding="utf-8") as f:
    f.write(app_code)

def run_streamlit():
    subprocess.run(["streamlit", "run", "app.py", "--server.port=8501"])

threading.Thread(target=run_streamlit, daemon=True).start()
time.sleep(3)

!curl -sL https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-linux-amd64 -o cloudflared
!chmod +x cloudflared

print("Đang tạo link truy cập mới...")
!./cloudflared tunnel --url http://localhost:8501