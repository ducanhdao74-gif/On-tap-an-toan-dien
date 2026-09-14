import json
import os
import random
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện", page_icon="⚡", layout="wide"
)

PROGRESS_FILE = "quiz_progress.json"


def load_saved_progress():
  if os.path.exists(PROGRESS_FILE):
    try:
      with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except:
      pass
  return {
      "completed_chunks": [0],
      "passed_tests": [0],
      "bookmarked_questions": [],
      "wrong_questions": [],
  }


def save_current_progress():
  data = {
      "completed_chunks": list(st.session_state.get("completed_chunks", [])),
      "passed_tests": list(st.session_state.get("passed_tests", [])),
      "bookmarked_questions": st.session_state.get("bookmarked_questions", []),
      "wrong_questions": st.session_state.get("wrong_questions", []),
  }
  try:
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
      json.dump(data, f, ensure_ascii=False, indent=4)
  except:
    pass


st.markdown(
    """
    <style>
    div[data-testid="stHorizontalBlock"] div.stRadio [role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        align-items: stretch !important;
    }
    div[data-testid="stHorizontalBlock"] div.stRadio [role="radiogroup"] label {
        text-align: left !important;
        font-size: 1rem !important;
    }

    .question-box {
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0px !important;
        margin-bottom: 15px;
    }
    .question-box h4 {
        font-size: 1.2rem !important;
        line-height: 1.6;
        color: #f8fafc !important;
        font-weight: 600;
    }
    
    div[data-testid="stRadio"] > div[role="radiogroup"] label p,
    div[data-testid="stRadio"] > div[role="radiogroup"] label span,
    div[data-testid="stRadio"] > div[role="radiogroup"] label div {
        font-size: 19px !important;
        line-height: 1.5 !important;
        color: #f1f5f9 !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


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
    import openpyxl

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


def get_shuffled_options(q_item, session_key):
  if session_key not in st.session_state:
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

    st.session_state[session_key] = {
        "options": shuffled_options,
        "correct": new_correct_letter,
    }

  return st.session_state[session_key]


sheets_data, error_message = load_data()
saved_prog = load_saved_progress()

if "wrong_questions" not in st.session_state:
  st.session_state["wrong_questions"] = saved_prog.get("wrong_questions", [])
if "bookmarked_questions" not in st.session_state:
  st.session_state["bookmarked_questions"] = saved_prog.get(
      "bookmarked_questions", []
  )
if "completed_chunks" not in st.session_state:
  st.session_state["completed_chunks"] = set(
      saved_prog.get("completed_chunks", [])
  )
if "passed_tests" not in st.session_state:
  st.session_state["passed_tests"] = set(saved_prog.get("passed_tests", []))

save_current_progress()

st.sidebar.title("⚡ Menu Ôn Tập")
st.sidebar.markdown("**Chọn chế độ:**")
mode = st.sidebar.radio(
    "",
    [
        "📖 Ôn tập theo chuyên đề",
        "🔄 Ôn lại câu trả lời sai",
        "📂 Ôn gộp tất cả (50 câu/phần)",
        "⭐ Tất cả câu hỏi cần ghi nhớ",
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

    if st.sidebar.button("🔄 Đặt lại tiến độ & Xáo trộn lại chuyên đề"):
      st.session_state[f"q_idx_{selected_sheet}"] = 0
      st.session_state[f"done_{selected_sheet}"] = 0
      keys_to_del = [
          k
          for k in st.session_state.keys()
          if k.startswith(f"shuff_chuande_{selected_sheet}_")
          or k.startswith(f"user_ans_chuande_{selected_sheet}_")
      ]
      for k in keys_to_del:
        del st.session_state[k]
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

      is_bm = any(
          b.get("question") == q_item["question"]
          for b in st.session_state["bookmarked_questions"]
      )
      bm_label = (
          "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
      )
      if st.button(bm_label, key=f"bm_chuande_{idx}"):
        if is_bm:
          st.session_state["bookmarked_questions"] = [
              b
              for b in st.session_state["bookmarked_questions"]
              if b.get("question") != q_item["question"]
          ]
          st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
        else:
          st.session_state["bookmarked_questions"].append(q_item)
          st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
        save_current_progress()
        st.rerun()

      shuff_data = get_shuffled_options(
          q_item, f"shuff_chuande_{selected_sheet}_{idx}"
      )
      options = shuff_data["options"]
      correct_letter = shuff_data["correct"]

      ans_storage_key = f"user_ans_chuande_{selected_sheet}_{idx}"
      default_idx = None
      current_saved_ans = st.session_state.get(ans_storage_key, None)
      if current_saved_ans in options:
        default_idx = options.index(current_saved_ans)

      selected_opt = st.radio(
          "Chọn đáp án của bạn:",
          options,
          index=default_idx,
          key=f"radio_chuande_{selected_sheet}_{idx}",
      )

      if selected_opt is not None:
        st.session_state[ans_storage_key] = selected_opt
        is_correct = selected_opt.strip().upper().startswith(correct_letter)
        if is_correct:
          st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
        else:
          st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
          if not any(
              w.get("question") == q_item["question"]
              for w in st.session_state["wrong_questions"]
          ):
            st.session_state["wrong_questions"].append(q_item)
            save_current_progress()

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
          f"Nguồn: {w_item.get('sheet', 'N/A')} (Câu {w_idx + 1}/{len(wrong_list)})"
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

      shuff_data = get_shuffled_options(w_item, f"shuff_wrong_{w_idx}")
      options = shuff_data["options"]
      correct_letter = shuff_data["correct"]

      w_storage_key = f"user_ans_wrong_{w_idx}"
      w_default_idx = None
      if st.session_state.get(w_storage_key) in options:
        w_default_idx = options.index(st.session_state.get(w_storage_key))

      w_choice = st.radio(
          "Chọn đáp án của bạn:",
          options,
          index=w_default_idx,
          key=f"radio_wrong_{w_idx}",
      )
      if w_choice is not None:
        st.session_state[w_storage_key] = w_choice
        if w_choice.strip().upper().startswith(correct_letter):
          st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
          wrong_list = [
              w
              for w in wrong_list
              if w.get("question") != w_item.get("question")
          ]
          st.session_state["wrong_questions"] = wrong_list
          save_current_progress()
        else:
          st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")

      st.markdown("---")
      if st.button("Câu tiếp theo ➡️", type="primary", key=f"next_wrong_{w_idx}"):
        if st.session_state["wrong_idx"] < len(wrong_list) - 1:
          st.session_state["wrong_idx"] += 1
        else:
          st.session_state["wrong_idx"] = 0
        st.rerun()

  elif mode == "⭐ Tất cả câu hỏi cần ghi nhớ":
    st.title("⭐ Tất Cả Câu Hỏi Cần Ghi Nhớ (Toàn Bộ Ngân Hàng)")
    st.markdown("---")

    bm_list = st.session_state["bookmarked_questions"]
    if not bm_list:
      st.info(
          "⭐ Bạn chưa đánh dấu câu hỏi nào cần ghi nhớ cả. Hãy bấm 'Đánh dấu"
          " câu cần ghi nhớ' trong lúc làm bài để lưu lại nhé!"
      )
    else:
      st.write(
          f"Tổng số câu bạn đã đánh dấu ghi nhớ: **{len(bm_list)}** câu."
      )
      if "global_bm_idx" not in st.session_state:
        st.session_state["global_bm_idx"] = 0

      gbm_idx = st.session_state["global_bm_idx"]
      if gbm_idx >= len(bm_list):
        gbm_idx = 0
        st.session_state["global_bm_idx"] = 0

      bm_item = bm_list[gbm_idx]
      st.subheader(
          f"Nguồn chuyên đề: {bm_item.get('sheet', 'N/A')} (Câu đánh dấu"
          f" {gbm_idx + 1}/{len(bm_list)})"
      )
      st.markdown("---")

      st.markdown(
          f"""
                <div class="question-box">
                    <h4>{bm_item['question']}</h4>
                </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("❌ Bỏ đánh dấu câu này", key=f"remove_global_bm_{gbm_idx}__"):
        st.session_state["bookmarked_questions"] = [
            b
            for b in st.session_state["bookmarked_questions"]
            if b.get("question") != bm_item.get("question")
        ]
        save_current_progress()
        st.toast("Đã xóa khỏi danh sách ghi nhớ!", icon="ℹ️")
        st.rerun()

      options = bm_item["options"]
      if not options:
        options = ["A. Đang cập nhật", "B. ---", "C. ---", "D. ---"]

      correct_letter = bm_item.get("correct", "A").strip().upper()

      st.markdown("##### 💡 Đáp án chuẩn:")
      for opt in options:
        opt_letter = opt.strip().upper()
        if opt_letter.startswith(correct_letter):
          st.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
        else:
          st.markdown(f"- {opt}")

      st.markdown("---")
      if st.button(
          "Câu tiếp theo ➡️", type="primary", key=f"next_global_bm_{gbm_idx}"
      ):
        if st.session_state["global_bm_idx"] < len(bm_list) - 1:
          st.session_state["global_bm_idx"] += 1
        else:
          st.session_state["global_bm_idx"] = 0
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
        is_done = (
            i in st.session_state["completed_chunks"]
            or i in st.session_state["passed_tests"]
        )
        prefix = "✅ [ĐÃ HOÀN THÀNH] " if is_done else ""
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
              "⚠️ Làm lại các câu sai trong phầnนี้",
          ],
          horizontal=True,
          label_visibility="collapsed",
      )
      st.markdown("---")

      if sub_mode == "📖 Ôn tập từng câu trong phần":
        if f"gop_idx_{c_num}" not in st.session_state:
          st.session_state[f"gop_idx_{c_num}"] = 0

        g_idx = st.session_state[f"gop_idx_{c_num}"]

        # KIỂM TRA NẾU ĐÃ LÀM XONG HẾT CÁC CÂU TRONG PHẦN
        if g_idx >= actual_chunk_len:
          st.success(
              f"🎉 Bạn đã hoàn thành phần ôn tập! (Phần {c_num + 1} - Tổng số"
              f" {actual_chunk_len} câu)"
          )
          st.markdown("---")
          col_btn1, col_btn2 = st.columns(2)
          with col_btn1:
            if st.button("🔄 Ôn lại (Làm lại từ đầu)", type="secondary"):
              st.session_state[f"gop_idx_{c_num}"] = 0
              keys_to_del = [
                  k
                  for k in st.session_state.keys()
                  if k.startswith(f"shuff_gop_{c_num}_")
                  or k.startswith(f"user_ans_gop_{c_num}_")
              ]
              for k in keys_to_del:
                del st.session_state[k]
              st.rerun()
          with col_btn2:
            if st.button(
                "📝 Bước đến bài thi (Kiểm tra chốt kiến thức)", type="primary"
            ):
              # Chuyển trực tiếp sang sub_mode bài kiểm tra
              # Streamlit không cho thay đổi radio index trực tiếp dễ dàng ngoài việc dùng session_state cho selectbox/radio hoặc hiển thị trực tiếp.
              # Cách tốt nhất là gán cờ hoặc ta tự động chuyển hướng hiển thị phần kiểm tra luôn ở đây.
              st.session_state[f"auto_switch_test_{c_num}"] = True
              st.rerun()
        else:
          # Kiểm tra nếu cờ auto_switch_test đang bật thì nhảy thẳng vào bài kiểm tra
          if st.session_state.get(f"auto_switch_test_{c_num}", False):
            # Reset cờ và chuyển sang tab kiểm tra bằng cách set lại sub_mode giả định
            st.session_state[f"auto_switch_test_{c_num}"] = False
            # Ta có thể render thẳng nội dung bài kiểm tra hoặc hướng dẫn người dùng bấm tab bài kiểm tra
            # Hoặc gán lại session_state để render giao diện kiểm tra luôn:
            sub_mode = "📝 Bài kiểm tra chốt kiến thức phần này"

        # Nếu sub_mode vẫn là ôn tập từng câu và g_idx < actual_chunk_len
        if sub_mode == "📖 Ôn tập từng câu trong phần" and g_idx < actual_chunk_len:
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

          is_bm = any(
              b.get("question") == q_item["question"]
              for b in st.session_state["bookmarked_questions"]
          )
          bm_label = (
              "⭐ Đã đánh dấu ghi nhớ" if is_bm else "☆ Đánh dấu câu cần ghi nhớ"
          )
          if st.button(bm_label, key=f"bm_gop_{c_num}_{g_idx}"):
            if is_bm:
              st.session_state["bookmarked_questions"] = [
                  b
                  for b in st.session_state["bookmarked_questions"]
                  if b.get("question") != q_item["question"]
              ]
              st.toast("Đã bỏ đánh dấu câu hỏi!", icon="ℹ️")
            else:
              st.session_state["bookmarked_questions"].append(q_item)
              st.toast("Đã thêm vào danh sách cần ghi nhớ!", icon="⭐")
            save_current_progress()
            st.rerun()

          shuff_data = get_shuffled_options(q_item, f"shuff_gop_{c_num}_{g_idx}")
          options = shuff_data["options"]
          correct_letter = shuff_data["correct"]

          gop_storage_key = f"user_ans_gop_{c_num}_{g_idx}"
          gop_default_idx = None
          if st.session_state.get(gop_storage_key) in options:
            gop_default_idx = options.index(
                st.session_state.get(gop_storage_key)
            )

          g_choice = st.radio(
              "Chọn đáp án của bạn:",
              options,
              index=gop_default_idx,
              key=f"radio_gop_{c_num}_{g_idx}",
          )

          if g_choice is not None:
            st.session_state[gop_storage_key] = g_choice
            is_correct = g_choice.strip().upper().startswith(correct_letter)
            if is_correct:
              st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
            else:
              st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")
              if not any(
                  w.get("question") == q_item["question"]
                  for w in st.session_state["wrong_questions"]
              ):
                st.session_state["wrong_questions"].append(q_item)
                save_current_progress()

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
              st.session_state[f"gop_idx_{c_num}"] += 1
              st.rerun()

      if sub_mode == "📝 Bài kiểm tra chốt kiến thức phần này":
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
            shuff_data = get_shuffled_options(q, f"shuff_test_{c_num}_{i}")
            options = shuff_data["options"]

            test_ans_key = f"test_chunk_{c_num}_{i}"
            ans = st.radio(
                "Chọn đáp án:",
                options,
                index=(
                    options.index(st.session_state[test_ans_key])
                    if st.session_state.get(test_ans_key) in options
                    else None
                ),
                key=test_ans_key,
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
                shuff_data = st.session_state.get(
                    f"shuff_test_{c_num}_{i}", {"correct": "A"}
                )
                correct_letter = shuff_data["correct"]

                if selected and selected.strip().upper().startswith(
                    correct_letter
                ):
                  correct_count += 1
                else:
                  wrong_count += 1
                  if not any(
                      w.get("question") == q["question"]
                      for w in st.session_state["wrong_questions"]
                  ):
                    st.session_state["wrong_questions"].append(q)

              st.session_state[f"result_correct_{c_num}"] = correct_count
              st.session_state[f"result_wrong_{c_num}"] = wrong_count
              st.session_state[f"test_user_answers_{c_num}"] = user_answers
              st.session_state[submitted_key] = True
              st.session_state["completed_chunks"].add(c_num)
              st.session_state["passed_tests"].add(c_num)
              save_current_progress()
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
          st.markdown("### 🔍 Xem Lại Chi Tiết Các Câu Trả Lời Sai")
          user_answers = st.session_state.get(f"test_user_answers_{c_num}", {})
          wrong_items_in_test = []
          for i, q in enumerate(current_chunk_questions):
            selected = user_answers.get(i)
            shuff_data = st.session_state.get(
                f"shuff_test_{c_num}_{i}", {"correct": "A"}
            )
            correct_letter = shuff_data["correct"]

            is_correct = selected and selected.strip().upper().startswith(
                correct_letter
            )
            if not is_correct:
              wrong_items_in_test.append((i, q, selected, shuff_data))

          if not wrong_items_in_test:
            st.info(
                "🎉 Tuyệt vời! Ông đã trả lời đúng tất cả các câu trong phần"
                " này."
            )
          else:
            st.write(
                f"Ông trả lời sai **{len(wrong_items_in_test)}** câu. Dưới đây"
                " là chi tiết các câu sai, đáp án ông đã chọn và đáp án đúng"
                " chuẩn:"
            )
            for q_idx, q_item, user_sel, shuff_info in wrong_items_in_test:
              st.markdown(
                  f"**Câu {q_idx + 1}** *(Thuộc chuyên đề: {q_item['sheet']})*"
              )
              st.markdown(f"> **{q_item['question']}**")

              correct_letter = shuff_info["correct"]

              for opt in shuff_info["options"]:
                opt_letter = opt.strip().upper()[:1]
                is_this_correct = opt_letter == correct_letter
                is_user_chosen = user_sel and opt.strip() == user_sel.strip()

                if is_this_correct:
                  st.markdown(f"- ✅ **{opt}** *(Đáp án đúng)*")
                elif is_user_chosen:
                  st.markdown(f"- ❌ ~~{opt}~~ *(Ông đã chọn)*")
                else:
                  st.markdown(f"- {opt}")

              if not user_sel:
                st.markdown("👉 *Ông chưa chọn đáp án nào cho câu này.*")

              st.markdown("---")

          if st.button("🔄 Làm lại bài kiểm tra này"):
            st.session_state[submitted_key] = False
            keys_to_del = [
                k
                for k in st.session_state.keys()
                if k.startswith(f"shuff_test_{c_num}_")
                or k.startswith(f"test_chunk_{c_num}_")
            ]
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

      elif sub_mode == "🔄 Làm lại phần này (Ôn tập lại từ đầu)":
        st.session_state[f"gop_idx_{c_num}"] = 0
        keys_to_del = [
            k
            for k in st.session_state.keys()
            if k.startswith(f"shuff_gop_{c_num}_")
            or k.startswith(f"shuff_test_{c_num}_")
            or k.startswith(f"user_ans_gop_{c_num}_")
            or k.startswith(f"test_chunk_{c_num}_")
        ]
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
        st.success(
            f"Đã reset và xáo trộn lại vị trí đáp án của Phần {c_num + 1}"
            " thành công!"
        )
        st.rerun()

      elif sub_mode == "⚠️ Làm lại các câu sai trong phần này":
        st.markdown(f"### Ôn lại các câu sai trong Phần {c_num + 1}")
        chunk_questions_texts = set(q["question"] for q in current_chunk_questions)
        wrong_in_chunk = [
            q
            for q in st.session_state["wrong_questions"]
            if q.get("question") in chunk_questions_texts
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

          shuff_data = get_shuffled_options(
              wc_item, f"shuff_wc_{c_num}_{wc_idx}"
          )
          options = shuff_data["options"]
          correct_letter = shuff_data["correct"]

          wc_storage_key = f"user_ans_wc_{c_num}_{wc_idx}"
          wc_default_idx = None
          if st.session_state.get(wc_storage_key) in options:
            wc_default_idx = options.index(st.session_state.get(wc_storage_key))

          wc_choice = st.radio(
              "Chọn đáp án của bạn:",
              options,
              index=wc_default_idx,
              key=f"radio_wc_{c_num}_{wc_idx}",
          )
          if wc_choice is not None:
            st.session_state[wc_storage_key] = wc_choice
            if wc_choice.strip().upper().startswith(correct_letter):
              st.success(f"🎉 Chính xác! Đáp án đúng là {correct_letter}.")
              st.session_state["wrong_questions"] = [
                  w
                  for w in st.session_state["wrong_questions"]
                  if w.get("question") != wc_item.get("question")
              ]
              save_current_progress()
            else:
              st.error(f"❌ Sai rồi! Đáp án đúng là {correct_letter}.")

          st.markdown("---")
          if st.button(
              "Câu tiếp theo ➡️", type="primary", key=f"next_wc_{c_num}_{wc_idx}"
          ):
            if (
                st.session_state[f"wrong_chunk_idx_{c_num}"]
                < len(wrong_in_chunk) - 1
            ):
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
        keys_to_del = [
            k
            for k in st.session_state.keys()
            if k.startswith("shuff_mock_") or k.startswith("mock_q_")
        ]
        for k in keys_to_del:
          del st.session_state[k]
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
        shuff_data = get_shuffled_options(q, f"shuff_mock_{i}")
        options = shuff_data["options"]

        mock_ans_key = f"mock_q_{i}"
        ans = st.radio(
            "Chọn đáp án:",
            options,
            index=(
                options.index(st.session_state[mock_ans_key])
                if st.session_state.get(mock_ans_key) in options
                else None
            ),
            key=mock_ans_key,
        )
        st.session_state["mock_answers"][i] = ans
        st.markdown("---")

      if st.button("📤 Nộp bài thi", type="primary"):
        st.success("Đã nộp bài thành công!")
        if st.button("Làm bài thi mới"):
          st.session_state["mock_started"] = False
          st.rerun()
