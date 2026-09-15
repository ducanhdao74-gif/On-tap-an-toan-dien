import os
import time
from datetime import datetime
import pandas as pd
import streamlit as st

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Ôn Tập An Toàn Điện",
    page_icon="⚡",
    layout="centered",
)

# --- KHỞI TẠO SESSION STATE ---
if "start_time" not in st.session_state:
  st.session_state.start_time = time.time()

if "completed_questions" not in st.session_state:
  st.session_state.completed_questions = set()


# --- HÀM TÍNH TỔNG THỜI GIAN ÔN TẬP ---
def get_total_study_time():
  current_session_duration = time.time() - st.session_state.start_time
  total_seconds = (
      st.session_state.get("accumulated_time", 0) + current_session_duration
  )
  gio = int(total_seconds // 3600)
  phut = int((total_seconds % 3600) // 60)
  if gio > 0:
    return f"{gio} giờ {phut} phút"
  else:
    return f"{phut} phút"


# --- HÀM TẠO NỘI DUNG THÔNG BÁO TIẾN ĐỘ MỚI ---
def get_progress_email_content():
  so_cau_hoan_thanh = len(st.session_state.completed_questions)
  tong_so_cau = st.session_state.get("total_questions_count", 0)
  str_thoi_gian = get_total_study_time()

  noi_dung = f"""
Chào Đức Anh,

Hôm nay là một ngày mới rồi! Hãy dành ra chút thời gian để vào ôn tập ngân hàng câu hỏi An Toàn Điện nhé:

📊 Tiến độ hiện tại của ông:
- Số câu đã hoàn thành: {so_cau_hoan_thanh}/{tong_so_cau} câu
- Tổng thời gian đã ôn tập: {str_thoi_gian}

Chúc ông ôn thi thật tốt và đạt kết quả cao!
"""
  return noi_dung


# --- TÌM VÀ ĐỌC FILE EXCEL CŨ CỦA ÔNG ---
excel_file = None
for f in os.listdir("."):
  if f.endswith(".xlsx") and (
      f.startswith("PL1") or "toan" in f.lower() or "cau" in f.lower()
  ):
    excel_file = f
    break

if not excel_file:
  for f in os.listdir("."):
    if f.endswith(".xlsx"):
      excel_file = f
      break

if excel_file and os.path.exists(excel_file):
  df = pd.read_excel(excel_file)
  st.session_state.total_questions_count = len(df)
else:
  df = None
  st.session_state.total_questions_count = 0

# --- GIAO DIỆN CHÍNH (GIỮ NGUYÊN BỐ CỤC CŨ) ---
st.title("⚡ Ôn Tập Ngân Hàng Câu Hỏi An Toàn Điện")
st.write("Chào Đức Anh! Giao diện ôn tập quen thuộc của ông đã sẵn sàng.")

if df is not None:
  # Hiển thị số liệu tổng quan
  col1, col2 = st.columns(2)
  with col1:
    st.metric(
        "Tiến độ hoàn thành",
        f"{len(st.session_state.completed_questions)} /"
        f" {st.session_state.total_questions_count} câu",
    )
  with col2:
    st.metric("Tổng thời gian ôn", get_total_study_time())

  st.divider()

  # Khu vực làm bài cũ của ông (được tích hợp lại đầy đủ)
  st.subheader("📝 Danh sách và Nội dung Ôn Tập")

  # Tab hoặc bộ lọc chọn chế độ làm bài như cũ
  che_do = st.radio(
      "Chọn chế độ:", ["Làm từng câu hỏi", "Xem toàn bộ danh sách"]
  )

  if che_do == "Làm từng câu hỏi":
    cau_chon = st.number_input(
        "Chọn câu số:", min_value=1, max_value=len(df), value=1, step=1
    )
    idx = cau_chon - 1
    row = df.iloc[idx]

    st.markdown(f"**Câu {cau_chon}:**")
    # Hiển thị thông tin hàng dữ liệu câu hỏi từ Excel
    st.write(row.to_dict())

    if st.button("Đã hoàn thành câu này"):
      st.session_state.completed_questions.add(cau_chon)
      st.success(f"Đã lưu tiến độ câu {cau_chon}!")
  else:
    st.dataframe(df)

  st.divider()

  # Phần xem thử nội dung thông báo gửi cho ông
  with st.expander("📩 Xem trước thông báo tiến độ"):
    st.code(get_progress_email_content(), language="text")

else:
  st.warning(
      "⚠️ Không tìm thấy tệp Excel ngân hàng câu hỏi. Vui lòng kiểm tra lại"
      " kho lưu trữ."
  )
