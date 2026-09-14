import os
import random
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện", page_icon="⚡", layout="wide"
)

st.markdown(
    """
    <style>
    /* Bỏ khung phần câu hỏi */
    .question-box {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0px !important;
        margin-bottom: 15px;
    }
    .question-box h4 {
        font-size: 1.35rem !important;
        line-height: 1.6;
        color: #f8fafc !important;
        font-weight: 600;
    }
    
    /* Bỏ khung đáp án, căn giữa và tăng kích thước chữ */
    div.stRadio > label {
        font-size: 1.1rem !important;
        font-weight: bold;
        color: #38bdf8 !important;
        margin-bottom: 10px;
    }
    div.stRadio [role="radiogroup"] {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0px !important;
        display: flex;
        flex-direction: column;
        align-items: center; /* Căn giữa các lựa chọn đáp án */
    }
    div.stRadio [role="radiogroup"] label {
        font-size: 1.15rem !important; /* Tăng chữ to lên xíu */
        padding: 6px 0;
        color: #e2e8f0 !important;
        text-align: center !important;
        width: 100%;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data():
  file_name = "PL1. Tong hop ngan hang cau hoi an toan nam 2025 fn (1).xlsx"
  if not os.path.exists(file_name):
    for f in os.listdir("."):
      if f.endswith(".xlsx") and "ngan hang cau hoi" in f.lower():
        file_name = f
        break
  if not os.path.exists(file_name):
    return None, f"Không tìm thấy file Excel trong thư mục!"
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
            questions.append({
                "question": current_q,
                "options": current_opts,
                "sheet": sheet,
            })
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
        questions.append({
            "question": current_q,
            "options": current_opts,
            "sheet": sheet,
        })

      if not questions:
        for idx, row in df.iterrows():
          row_vals = [str(x) for x in row.values if pd.notna(x)]
          if row_vals:
            questions.append({
                "question": row_vals[0],
                "options": (
                    row_vals[1:]
                    if len(row_vals) > 1
                    else ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
                ),
                "sheet": sheet,
            })

      sheets_data[sheet] = questions
    return sheets_data, None
  except Exception as e:
    return None, f"Lỗi khi đọc file Excel: {e}"


sheets_data, error_message = load_data()

if "wrong_questions" not in st.session_state:
  st.session_state["wrong_questions"] = []
if "bookmarked_questions" not in st.session_state:
  st.session_state["bookmarked_questions"] = []
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
        "📝 Thi thử (Mock Test)",
    ],
    label_visibility="collapsed",
)

if error_message:
  st.error(error_message)
else:
  if mode == "📖 Ôn tập theo chuyên đề":
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📂 Chọn Chuyên Đề")
    sheet_list = list(sheets_data.keys())
    selected_sheet = st.sidebar.selectbox(
        "", sheet_list, label_visibility="collapsed"
    )

    q_list = sheets_data[selected_sheet]
    total_q = len(q_list)

    if f"q_idx_{selected_sheet}" not in st.session_state:
      st.session_state[f"q_idx_{selected_sheet}"] = 0
    if f"done_{selected_sheet}" not in st.session_state:
      st.session_state[f"done_{selected_sheet}"] = 0

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**Tổng số câu**\n### {total_q}")
    st.sidebar.markdown(
        f"**Đã làm**\n### {st.session_state[f'done_{selected_sheet}']}"
    )

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
      st.subheader(
          f"Chuyên đề: {selected_sheet} (Câu {idx + 1}/{total_q})"
      )
      st.markdown("---")

      st.markdown(
          f"""
                <div class="question-box">
                    <h4>{q_item['question']}</h4>
                </div>
            """,
          unsafe_allow_html=True,
      )

      is_bm = q_item in st.session_state["bookmarked_questions"]
      bm_label = (
          "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
      )
      if st.button(bm_label, key=f"bm_chuande_{idx}"):
        if is_bm:
          st.session_state["bookmarked_questions"].remove(q_item)
          st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
        else:
          st.session_state["bookmarked_questions"].append(q_item)
          st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
        st.rerun()

      options = q_item["options"]
      if not options:
        options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

      ans_key = f"ans_chuande_{selected_sheet}_{idx}"
      selected_opt = st.radio(
          "Chọn đáp án của bạn:", options, index=None, key=ans_key
      )

      if selected_opt is not None:
        is_correct = selected_opt.strip().startswith("A.")
        if is_correct:
          st.success("🎉 Chính xác! Đáp án đúng là A.")
        else:
          st.error("❌ Sai rồi! Đáp án đúng là A.")
          if q_item not in st.session_state["wrong_questions"]:
            st.session_state["wrong_questions"].append(q_item)

        st.session_state[f"done_{selected_sheet}"] = min(
            total_q,
            max(st.session_state[f"done_{selected_sheet}"], idx + 1),
        )

      st.markdown("---")
      if st.button(
          "Câu tiếp theo ➡️", type="primary", key=f"next_chuande_{idx}"
      ):
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
      st.info(
          "🎉 Hiện tại bạn chưa có câu trả lời sai nào được lưu lại. Hãy làm"
          " bài để hệ thống ghi nhận nhé!"
      )
    else:
      st.write(f"Bạn đang có **{len(wrong_list)}** câu cần ôn tập lại.")
      if "wrong_idx" not in st.session_state:
        st.session_state["wrong_idx"] = 0

      w_idx = st.session_state["wrong_idx"]
      if w_idx >= len(wrong_list):
        w_idx = 0
        st.session_state["wrong_idx"] = 0

      w_item = wrong_list[w_idx]
      st.subheader(
          f"Nguồn: {w_item['sheet']} (Câu {w_idx + 1}/{len(wrong_list)})"
      )
      st.markdown("---")

      st.markdown(
          f"""
                <div class="question-box">
                    <h4>{w_item['question']}</h4>
                </div>
            """,
          unsafe_allow_html=True,
      )

      w_choice = st.radio(
          "Chọn đáp án của bạn:",
          w_item["options"],
          index=None,
          key=f"radio_wrong_{w_idx}",
      )
      if w_choice is not None:
        if w_choice.strip().startswith("A."):
          st.success("🎉 Chính xác! Đáp án đúng là A.")
          if w_item in wrong_list:
            wrong_list.remove(w_item)
            st.session_state["wrong_questions"] = wrong_list
        else:
          st.error("❌ Sai rồi! Đáp án đúng là A.")

      st.markdown("---")
      if st.button("Câu tiếp theo ➡️", type="primary", key=f"next_wrong_{w_idx}"):
        if st.session_state["wrong_idx"] < len(wrong_list) - 1:
          st.session_state["wrong_idx"] += 1
        else:
          st.session_state["wrong_idx"] = 0
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
      total_chunks = (total_all // chunk_size) + (
          1 if total_all % chunk_size != 0 else 0
      )

      st.sidebar.markdown("---")
      st.sidebar.markdown("### 📂 Chọn Phần Ôn Tập")

      chunk_names = []
      for i in range(total_chunks):
        prefix = (
            "✅ [ĐÃ HOÀN THÀNH] "
            if i in st.session_state["completed_chunks"]
            else ""
        )
        chunk_names.append(f"{prefix}Phần {i+1} (Hỗn hợp chuyên đề)")

      selected_chunk_name = st.sidebar.selectbox(
          "", chunk_names, label_visibility="collapsed"
      )
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

      sub_mode = st.radio(
          "",
          [
              "📖 Ôn tập từng câu trong phần",
              "📝 Bài kiểm tra chốt kiến thức phần này",
              "🔄 Làm lại phần này (Ôn tập lại từ đầu)",
              "⚠️ Làm lại các câu sai trong phần này",
              "⭐ Câu hỏi cần ghi nhớ",
          ],
          horizontal=True,
          label_visibility="collapsed",
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

        st.markdown(
            f"### Đang ôn: Phần {c_num + 1} (Hỗn hợp chuyên đề) — *(Thuộc"
            f" chuyên đề: {q_item['sheet']})*"
        )
        st.markdown("---")

        st.markdown(
            f"""
                    <div class="question-box">
                        <h4>Câu {g_idx + 1}/{actual_chunk_len} (Toàn hệ thống #{start_idx + g_idx + 1}): {q_item['question']}</h4>
                    </div>
                """,
            unsafe_allow_html=True,
        )

        is_bm = q_item in st.session_state["bookmarked_questions"]
        bm_label = (
            "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
        )
        if st.button(bm_label, key=f"bm_gop_{c_num}_{g_idx}"):
          if is_bm:
            st.session_state["bookmarked_questions"].remove(q_item)
            st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
          else:
            st.session_state["bookmarked_questions"].append(q_item)
            st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
          st.rerun()

        options = q_item["options"]
        if not options:
          options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

        g_choice = st.radio(
            "Chọn đáp án của bạn:",
            options,
            index=None,
            key=f"radio_gop_{c_num}_{g_idx}",
        )

        if g_choice is not None:
          is_correct = g_choice.strip().startswith("A.")
          if is_correct:
            st.success("🎉 Chính xác! Đáp án đúng là A.")
          else:
            st.error("❌ Sai rồi! Đáp án đúng là A.")
            if q_item not in st.session_state["wrong_questions"]:
              st.session_state["wrong_questions"].append(q_item)

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
          if st.button("Câu tiếp theo ➡️", type="primary"):
            if st.session_state[f"gop_idx_{c_num}"] < actual_chunk_len - 1:
              st.session_state[f"gop_idx_{c_num}"] += 1
            else:
              st.session_state[f"gop_idx_{c_num}"] = 0
              st.session_state["completed_chunks"].add(c_num)
            st.rerun()

      elif sub_mode == "📝 Bài kiểm tra chốt kiến thức phần này":
        st.markdown(
            f"### Bài kiểm tra - Phần {c_num + 1} ({actual_chunk_len} câu)"
        )

        submitted_key = f"submitted_test_{c_num}"
        if submitted_key not in st.session_state:
          st.session_state[submitted_key] = False

        if not st.session_state[submitted_key]:
          user_answers = {}
          for i, q in enumerate(current_chunk_questions):
            st.markdown(
                f"""
                            <div class="question-box">
                                <h4>Câu {i+1} *(Thuộc: {q['sheet']})*: {q['question']}</h4>
                            </div>
                        """,
                unsafe_allow_html=True,
            )
            options = (
                q["options"]
                if q["options"]
                else ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
            )

            ans = st.radio(
                "Chọn đáp án:",
                options,
                index=None,
                key=f"test_chunk_{c_num}_{i}",
            )
            user_answers[i] = ans
            st.markdown("---")

          if st.button("Nộp bài kiểm tra", type="primary"):
            unanswered = [
                i + 1 for i in range(actual_chunk_len) if user_answers[i] is None
            ]
            if unanswered:
              st.error(
                  f"⚠️ Ông chưa làm xong tất cả các câu! Còn thiếu các câu:"
                  f" {', '.join(map(str, unanswered))}."
              )
            else:
              correct_count = 0
              wrong_count = 0
              for i, q in enumerate(current_chunk_questions):
                selected = user_answers[i]
                if selected and selected.strip().startswith("A."):
                  correct_count += 1
                else:
                  wrong_count += 1
                  if q not in st.session_state["wrong_questions"]:
                    st.session_state["wrong_questions"].append(q)

              st.session_state[f"result_correct_{c_num}"] = correct_count
              st.session_state[f"result_wrong_{c_num}"] = wrong_count
              st.session_state[submitted_key] = True
              st.session_state["completed_chunks"].add(c_num)
              st.rerun()
        else:
          c_correct = st.session_state.get(f"result_correct_{c_num}", 0)
          c_wrong = st.session_state.get(f"result_wrong_{c_num}", 0)
          score_percent = (c_correct / actual_chunk_len) * 100

          st.success("🎉 Đã nộp bài kiểm tra thành công!")
          st.markdown("### 📊 Kết Quả Bài Kiểm Tra")
          col1, col2, col3 = st.columns(3)
          col1.metric(
              "Số câu đúng",
              f"{c_correct}/{actual_chunk_len}",
              f"{score_percent:.1f}%",
          )
          col2.metric("Số câu sai", f"{c_wrong}/{actual_chunk_len}")
          col3.metric("Trạng thái", "Đã hoàn thành ✅")

          st.markdown("---")
          if st.button("🔄 Làm lại bài kiểm tra này"):
            st.session_state[submitted_key] = False
            st.rerun()

      elif sub_mode == "🔄 Làm lại phần này (Ôn tập lại từ đầu)":
        st.session_state[f"gop_idx_{c_num}"] = 0
        if f"submitted_test_{c_num}" in st.session_state:
          st.session_state[f"submitted_test_{c_num}"] = False
        if c_num in st.session_state["completed_chunks"]:
          st.session_state["completed_chunks"].remove(c_num)
        st.success(
            f"Đã reset và xóa trạng thái hoàn thành của Phần {c_num + 1}"
            " thành công!"
        )
        st.rerun()

      elif sub_mode == "⚠️ Làm lại các câu sai trong phần này":
        st.markdown(f"### Ôn lại các câu sai trong Phần {c_num + 1}")
        chunk_questions_set = set(id(q) for q in current_chunk_questions)
        wrong_in_chunk = [
            q
            for q in st.session_state["wrong_questions"]
            if id(q) in chunk_questions_set
        ]

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
          st.write(
              f"Đang ôn câu sai **{wc_idx + 1}/{len(wrong_in_chunk)}** trong"
              " phần này:"
          )
          st.markdown("---")

          st.markdown(
              f"""
                        <div class="question-box">
                            <h4>{wc_item['question']}</h4>
                        </div>
                    """,
              unsafe_allow_html=True,
          )

          wc_choice = st.radio(
              "Chọn đáp án của bạn:",
              wc_item["options"],
              index=None,
              key=f"radio_wc_{c_num}_{wc_idx}",
          )
          if wc_choice is not None:
            if wc_choice.strip().startswith("A."):
              st.success("🎉 Chính xác! Đáp án đúng là A.")
              if wc_item in st.session_state["wrong_questions"]:
                st.session_state["wrong_questions"].remove(wc_item)
            else:
              st.error("❌ Sai rồi! Đáp án đúng là A.")

          st.markdown("---")
          if st.button(
              "Câu tiếp theo ➡️", type="primary", key=f"next_wc_{c_num}_{wc_idx}"
          ):
            if st.session_state[f"wrong_chunk_idx_{c_num}"] < len(wrong_in_chunk) - 1:
              st.session_state[f"wrong_chunk_idx_{c_num}"] += 1
            else:
              st.session_state[f"wrong_chunk_idx_{c_num}"] = 0
            st.rerun()

      elif sub_mode == "⭐ Câu hỏi cần ghi nhớ":
        st.markdown(
            f"### ⭐ Danh Sách Câu Hỏi Cần Ghi Nhớ (Phần {c_num + 1})"
        )
        chunk_questions_set = set(id(q) for q in current_chunk_questions)
        bm_in_chunk = [
            q
            for q in st.session_state["bookmarked_questions"]
            if id(q) in chunk_questions_set
        ]

        if not bm_in_chunk:
          st.info(
              "⭐ Phần này bạn chưa đánh dấu câu hỏi nào cần ghi nhớ cả."
          )
        else:
          if f"bm_chunk_idx_{c_num}" not in st.session_state:
            st.session_state[f"bm_chunk_idx_{c_num}"] = 0

          bmc_idx = st.session_state[f"bm_chunk_idx_{c_num}"]
          if bmc_idx >= len(bm_in_chunk):
            bmc_idx = 0
            st.session_state[f"bm_chunk_idx_{c_num}"] = 0

          bmc_item = bm_in_chunk[bmc_idx]
          st.write(
              f"Đang xem câu ghi nhớ **{bmc_idx + 1}/{len(bm_in_chunk)}** trong"
              " phần này:"
          )
          st.markdown("---")

          st.markdown(
              f"""
                        <div class="question-box">
                            <h4>{bmc_item['question']}</h4>
                        </div>
                    """,
              unsafe_allow_html=True,
          )

          if st.button("❌ Bỏ đánh dấu câu này", key=f"remove_bm_{c_num}_{bmc_idx}"):
            st.session_state["bookmarked_questions"].remove(bmc_item)
            st.toast("Đã xóa khỏi danh sách ghi nhớ!", icon="ℹ️")
            st.rerun()

          bmc_choice = st.radio(
              "Chọn đáp án của bạn:",
              bmc_item["options"],
              index=None,
              key=f"radio_bm_{c_num}_{bmc_idx}",
          )
          if bmc_choice is not None:
            if bmc_choice.strip().startswith("A."):
              st.success("🎉 Chính xác! Đáp án đúng là A.")
            else:
              st.error("❌ Sai rồi! Đáp án đúng là A.")

          st.markdown("---")
          if st.button(
              "Câu tiếp theo ➡️", type="primary", key=f"next_bm_{c_num}_{bmc_idx}"
          ):
            if st.session_state[f"bm_chunk_idx_{c_num}"] < len(bm_in_chunk) - 1:
              st.session_state[f"bm_chunk_idx_{c_num}"] += 1
            else:
              st.session_state[f"bm_chunk_idx_{c_num}"] = 0
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
      st.write(
          "Đề thi thử gồm **50 câu hỏi ngẫu nhiên** được trộn từ toàn bộ ngân"
          " hàng câu hỏi an toàn điện."
      )
      if st.button("🚀 Bắt đầu làm bài thi", type="primary"):
        st.session_state["mock_started"] = True
        st.session_state["mock_questions"] = random.sample(
            all_questions, min(50, len(all_questions))
        )
        st.session_state["mock_answers"] = {}
        st.rerun()
    else:
      mock_qs = st.session_state["mock_questions"]
      st.write(f"Đang làm bài thi thử ({len(mock_qs)} câu).")

      for i, q in enumerate(mock_qs):
        st.markdown(
            f"""
                    <div class="question-box">
                        <h4>Câu {i+1}: {q['question']}</h4>
                    </div>
                """,
            unsafe_allow_html=True,
        )
        options = (
            q["options"]
            if q["options"]
            else ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]
        )

        ans = st.radio("Chọn đáp án:", options, index=None, key=f"mock_q_{i}")
        st.session_state["mock_answers"][i] = ans
        st.markdown("---")

      if st.button("📤 Nộp bài thi", type="primary"):
        st.success("Đã nộp bài thành công!")
        if st.button("Làm bài thi mới"):
          st.session_state["mock_started"] = False
          st.rerun()