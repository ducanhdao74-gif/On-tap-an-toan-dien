import streamlit as st
import pandas as pd
import os
import random

st.set_page_config(page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện", page_icon="⚡", layout="wide")

@st.cache_data
def load_data():
    file_name = "PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx"
    if not os.path.exists(file_name):
        return None, f"Không tìm thấy file `{file_name}` trong thư mục!"
    try:
        xls = pd.ExcelFile(file_name)
        sheets_data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            col = df.columns[0]
            
            questions = []
            current_q = None
            current_opts = []
            
            for val in df[col].dropna():
                val_str = str(val).strip()
                if val_str.lower().startswith("câu"):
                    if current_q:
                        questions.append({"question": current_q, "options": current_opts, "sheet": sheet})
                    current_q = val_str
                    current_opts = []
                elif val_str.lower().startswith(("a.", "b.", "c.", "d.")):
                    current_opts.append(val_str)
                else:
                    if current_q and not current_opts:
                        current_q += " " + val_str
                    elif current_opts:
                        current_opts[-1] += " " + val_str
            if current_q:
                questions.append({"question": current_q, "options": current_opts, "sheet": sheet})
                
            if not questions:
                for idx, row in df.iterrows():
                    row_vals = [str(x) for x in row.values if pd.notna(x)]
                    if row_vals:
                        questions.append({
                            "question": row_vals[0], 
                            "options": row_vals[1:] if len(row_vals)>1 else ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"],
                            "sheet": sheet
                        })
                        
            sheets_data[sheet] = questions
        return sheets_data, None
    except Exception as e:
        return None, f"Lỗi khi đọc file Excel: {e}"

sheets_data, error_message = load_data()

if "wrong_questions" not in st.session_state:
    st.session_state["wrong_questions"] = []
if "completed_chunks" not in st.session_state:
    st.session_state["completed_chunks"] = set()

st.sidebar.title("⚡ Menu Ôn Tập")
st.sidebar.markdown("**Chọn chế độ:**")
mode = st.sidebar.radio(
    "",
    [
        "📖 Ôn tập theo chuyên đề",
        "🔄 Ôn lại câu trả lời sai",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "📝 Thi thử (Mock Test)"
    ],
    label_visibility="collapsed"
)

if error_message:
    st.error(error_message)
else:
    if mode == "📖 Ôn tập theo chuyên đề":
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 📂 Chọn Chuyên Đề")
        sheet_list = list(sheets_data.keys())
        selected_sheet = st.sidebar.selectbox("", sheet_list, label_visibility="collapsed")
        
        q_list = sheets_data[selected_sheet]
        total_q = len(q_list)
        
        if f"q_idx_{selected_sheet}" not in st.session_state:
            st.session_state[f"q_idx_{selected_sheet}"] = 0
        if f"done_{selected_sheet}" not in st.session_state:
            st.session_state[f"done_{selected_sheet}"] = 0

        st.sidebar.markdown("---")
        st.sidebar.markdown(f"**Tổng số câu**\n### {total_q}")
        st.sidebar.markdown(f"**Đã làm**\n### {st.session_state[f'done_{selected_sheet}']}")
        
        if st.sidebar.button("🔄 Đặt lại tiến độ chuyên đề này"):
            st.session_state[f"q_idx_{selected_sheet}"] = 0
            st.session_state[f"done_{selected_sheet}"] = 0
            st.rerun()

        st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
        
        idx = st.session_state[f"q_idx_{selected_sheet}"]
        if idx >= total_q and total_q > 0:
            idx = 0
            st.session_state[f"q_idx_{selected_sheet}"] = 0

        if total_q == 0:
            st.warning("Chuyên đề này hiện chưa có câu hỏi nào.")
        else:
            q_item = q_list[idx]
            st.subheader(f"Chuyên đề: {selected_sheet} (Câu {idx + 1}/{total_q})")
            st.markdown("---")
            st.markdown(f"#### {q_item['question']}")
            st.write("Bấm chọn trực tiếp đáp án bên dưới:")

            options = q_item['options']
            if not options:
                options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

            for opt in options:
                if st.button(opt, key=f"btn_chuande_{idx}_{opt}", use_container_width=True):
                    is_correct = opt.strip().startswith("A.")
                    if is_correct:
                        st.toast("🎉 Chính xác!", icon="✅")
                    else:
                        st.toast("❌ Sai rồi! Đáp án đúng là A.", icon="⚠️")
                        if q_item not in st.session_state["wrong_questions"]:
                            st.session_state["wrong_questions"].append(q_item)
                    
                    st.session_state[f"done_{selected_sheet}"] = min(total_q, st.session_state[f"done_{selected_sheet}"] + 1)
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
            st.info("🎉 Hiện tại bạn chưa có câu trả lời sai nào được lưu lại. Hãy làm bài để hệ thống ghi nhận nhé!")
        else:
            st.write(f"Bạn đang có **{len(wrong_list)}** câu cần ôn tập lại.")
            if "wrong_idx" not in st.session_state:
                st.session_state["wrong_idx"] = 0
                
            w_idx = st.session_state["wrong_idx"]
            if w_idx >= len(wrong_list):
                w_idx = 0
                st.session_state["wrong_idx"] = 0
                
            w_item = wrong_list[w_idx]
            st.subheader(f"Nguồn: {w_item['sheet']} (Câu {w_idx + 1}/{len(wrong_list)})")
            st.markdown(f"#### {w_item['question']}")
            
            for opt in w_item['options']:
                if st.button(opt, key=f"btn_wrong_{w_idx}_{opt}", use_container_width=True):
                    if opt.strip().startswith("A."):
                        st.toast("🎉 Chính xác!", icon="✅")
                        wrong_list.pop(w_idx)
                        st.session_state["wrong_questions"] = wrong_list
                    else:
                        st.toast("❌ Vẫn chưa chính xác!", icon="⚠️")
                    st.rerun()

    elif mode == "📂 Ôn gộp tất cả (50 câu/phần)":
        st.title("📂 Ôn Gộp Tất Cả Chuyên Đề (Trộn Đều & Chia Phần 50 Câu)")
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
            st.sidebar.markdown("### 📂 Chọn Phần Ôn Tập")
            
            # Tạo tên hiển thị có gắn nhãn [ĐÃ HOÀN THÀNH] nếu đã xong
            chunk_names = []
            for i in range(total_chunks):
                prefix = "✅ [ĐÃ HOÀN THÀNH] " if i in st.session_state["completed_chunks"] else ""
                chunk_names.append(f"{prefix}Phần {i+1} (Hỗn hợp chuyên đề)")
                
            selected_chunk_name = st.sidebar.selectbox("", chunk_names, label_visibility="collapsed")
            c_num = chunk_names.index(selected_chunk_name)
            
            start_idx = c_num * chunk_size
            end_idx = min((c_num + 1) * chunk_size, total_all)
            
            if f"chunk_q_{c_num}" not in st.session_state:
                chunk_qs = all_questions.copy()
                random.seed(42)
                random.shuffle(chunk_qs)
                st.session_state[f"chunk_q_{c_num}"] = chunk_qs[start_idx:end_idx]
                
            current_chunk_questions = st.session_state[f"chunk_q_{c_num}"]
            actual_chunk_len = len(current_chunk_questions)
            
            st.sidebar.markdown(f"**Tổng số câu của phần**\n### {actual_chunk_len}")

            st.markdown("**Chọn chế độ cho phần này:**")
            sub_mode = st.radio(
                "",
                [
                    "📖 Ôn tập từng câu trong phần",
                    "📝 Bài kiểm tra chốt kiến thức phần này",
                    "🔄 Làm lại phần này (Ôn tập lại từ đầu)",
                    "⚠️ Làm lại các câu sai trong phần này"
                ],
                horizontal=True,
                label_visibility="collapsed"
            )
            st.markdown("---")

            if sub_mode == "📖 Ôn tập từng câu trong phần":
                if f"gop_idx_{c_num}" not in st.session_state:
                    st.session_state[f"gop_idx_{c_num}"] = 0
                    
                g_idx = st.session_state[f"gop_idx_{c_num}"]
                if g_idx >= actual_chunk_len:
                    g_idx = 0
                    st.session_state[f"gop_idx_{c_num}"] = 0
                    
                q_item = current_chunk_questions[g_idx]
                
                st.markdown(f"### Đang ôn: Phần {c_num + 1} (Hỗn hợp chuyên đề)")
                st.markdown(f"*(Thuộc chuyên đề: {q_item['sheet']})*")
                st.markdown(f"#### Câu {g_idx + 1}/{actual_chunk_len} (Toàn hệ thống #{start_idx + g_idx + 1}): {q_item['question']}")
                st.write("Bấm chọn trực tiếp đáp án bên dưới để tự động chuyển câu:")

                options = q_item['options']
                if not options:
                    options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

                for opt in options:
                    if st.button(opt, key=f"btn_gop_{c_num}_{g_idx}_{opt}", use_container_width=True):
                        is_correct = opt.strip().startswith("A.")
                        if is_correct:
                            st.toast("🎉 Chính xác!", icon="✅")
                        else:
                            st.toast("❌ Sai rồi! Đáp án đúng là A.", icon="⚠️")
                            if q_item not in st.session_state["wrong_questions"]:
                                st.session_state["wrong_questions"].append(q_item)

                        if g_idx < actual_chunk_len - 1:
                            st.session_state[f"gop_idx_{c_num}"] += 1
                        else:
                            st.session_state[f"gop_idx_{c_num}"] = 0
                            # Đánh dấu hoàn thành phần này khi đi hết vòng
                            st.session_state["completed_chunks"].add(c_num)
                        st.rerun()

                st.markdown("---")
                col_prev, col_next = st.columns(2)
                with col_prev:
                    if st.button("⬅️ Câu trước"):
                        if st.session_state[f"gop_idx_{c_num}"] > 0:
                            st.session_state[f"gop_idx_{c_num}"] -= 1
                        else:
                            st.session_state[f"gop_idx_{c_num}"] = actual_chunk_len - 1
                        st.rerun()
                with col_next:
                    if st.button("Câu tiếp theo ➡️"):
                        if st.session_state[f"gop_idx_{c_num}"] < actual_chunk_len - 1:
                            st.session_state[f"gop_idx_{c_num}"] += 1
                        else:
                            st.session_state[f"gop_idx_{c_num}"] = 0
                            st.session_state["completed_chunks"].add(c_num)
                        st.rerun()

            elif sub_mode == "📝 Bài kiểm tra chốt kiến thức phần này":
                st.markdown(f"### Bài kiểm tra - Phần {c_num + 1} ({actual_chunk_len} câu)")
                for i, q in enumerate(current_chunk_questions):
                    st.markdown(f"**Câu {i+1}** *(Thuộc: {q['sheet']})*: {q['question']}")
                    st.radio("Chọn đáp án:", q['options'], key=f"test_chunk_{c_num}_{i}")
                    st.markdown("---")
                if st.button("Nộp bài kiểm tra", type="primary"):
                    st.success("Đã hoàn thành bài kiểm tra phần này!")
                    st.session_state["completed_chunks"].add(c_num)
                    st.rerun()

            elif sub_mode == "🔄 Làm lại phần này (Ôn tập lại từ đầu)":
                st.session_state[f"gop_idx_{c_num}"] = 0
                if c_num in st.session_state["completed_chunks"]:
                    st.session_state["completed_chunks"].remove(c_num)
                st.success(f"Đã đặt lại tiến độ Phần {c_num + 1} về câu đầu tiên!")
                if st.button("Bắt đầu ôn ngay"):
                    st.rerun()

            elif sub_mode == "⚠️ Làm lại các câu sai trong phần này":
                st.markdown(f"### Ôn lại các câu sai trong Phần {c_num + 1}")
                chunk_questions_set = set(id(q) for q in current_chunk_questions)
                wrong_in_chunk = [q for q in st.session_state["wrong_questions"] if id(q) in chunk_questions_set]
                
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
                    st.write(f"Đang ôn câu sai **{wc_idx + 1}/{len(wrong_in_chunk)}** trong phần này:")
                    st.markdown(f"#### {wc_item['question']}")
                    
                    for opt in wc_item['options']:
                        if st.button(opt, key=f"btn_wrong_chunk_{c_num}_{wc_idx}_{opt}", use_container_width=True):
                            if opt.strip().startswith("A."):
                                st.toast("🎉 Chính xác!", icon="✅")
                                if wc_item in st.session_state["wrong_questions"]:
                                    st.session_state["wrong_questions"].remove(wc_item)
                            else:
                                st.toast("❌ Vẫn chưa chính xác!", icon="⚠️")
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
            st.write("Đề thi thử gồm **50 câu hỏi ngẫu nhiên** được trộn từ toàn bộ ngân hàng câu hỏi an toàn điện.")
            if st.button("🚀 Bắt đầu làm bài thi", type="primary"):
                st.session_state["mock_started"] = True
                st.session_state["mock_questions"] = random.sample(all_questions, min(50, len(all_questions)))
                st.session_state["mock_answers"] = {}
                st.rerun()
        else:
            mock_qs = st.session_state["mock_questions"]
            st.write(f"Đang làm bài thi thử ({len(mock_qs)} câu).")
            
            for i, q in enumerate(mock_qs):
                st.markdown(f"**Câu {i+1}: {q['question']}**")
                ans = st.radio("Chọn đáp án:", q['options'], key=f"mock_q_{i}")
                st.session_state["mock_answers"][i] = ans
                st.markdown("---")
                
            if st.button("📤 Nộp bài thi", type="primary"):
                st.success("Đã nộp bài thành công!")
                if st.button("Làm bài thi mới"):
                    st.session_state["mock_started"] = False
                    st.rerun()