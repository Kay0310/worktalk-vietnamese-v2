import streamlit as st
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io

st.markdown("<h1 style='text-align:center;'>WORK TALK</h1>", unsafe_allow_html=True)
st.markdown("<p style='font-size:16px; text-align:center;'>Tham gia hệ thống đánh giá tính nguy hiểm</p>", unsafe_allow_html=True)

st.markdown("<h3 style='margin-top: 20px;'>✏️ Thông tin người điền</h3>", unsafe_allow_html=True)
name = st.text_input("Vui lòng nhập tên người điền")
department = st.text_input("Vui lòng nhập bộ phận làm việc")

st.markdown("<h3 style='margin-top: 20px;'>📷 Tải ảnh công việc nguy hiểm</h3>", unsafe_allow_html=True)
uploaded_file = st.file_uploader("Tải ảnh công việc", type=['jpg', 'jpeg', 'png'])

if uploaded_file is not None:
    st.image(uploaded_file, caption="Xem trước ảnh đã tải lên", use_column_width=True)

st.markdown("<h3 style='margin-top: 20px;'>📋 Câu hỏi đánh giá rủi ro</h3>", unsafe_allow_html=True)
place = st.text_input("0. Công đoạn bạn đang thực hiện là gì?")
work = st.text_input("1. Bạn đang thực hiện công việc gì?")
danger_reason = st.text_input("2. Tại sao bạn nghĩ công việc này nguy hiểm?")

freq = st.radio("3. Công việc này được thực hiện thường xuyên như thế nào?", 
                ["1-2 lần/năm", "1-2 lần/6 tháng", "2-3 lần/tháng", "Ít nhất 1 lần/tuần", "Hàng ngày"])

risk = st.radio("4. Bạn đánh giá mức độ nguy hiểm của công việc này như thế nào?", 
                ["Ít nguy hiểm", "Hơi nguy hiểm", "Nguy hiểm", "Rất nguy hiểm"])

improvement = st.text_area("5. Nếu có ý tưởng cải thiện để công việc an toàn hơn, hãy ghi vào đây (Không bắt buộc)")

if st.button("Gửi bài"):
    if not name or not department or not uploaded_file:
        st.error("Tên, bộ phận và ảnh là bắt buộc!")
    else:
        st.success("Gửi bài thành công! Bây giờ bạn có thể tải xuống tệp Excel.")

        now = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"위험성평가_{name}_{now}.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = "Kết quả đánh giá"

        headers = ["Tên người điền", "Bộ phận làm việc", "Công đoạn", "Công việc thực hiện", 
                   "Lý do nguy hiểm", "Tần suất công việc", "Mức độ nguy hiểm", "Ý tưởng cải thiện"]
        values = [name, department, place, work, danger_reason, freq, risk, improvement]
        ws.append(headers)
        ws.append(values)

        img = Image.open(uploaded_file)
        img.thumbnail((150, 150))
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        img_for_excel = XLImage(img_byte_arr)
        ws.add_image(img_for_excel, 'I2')

        wb.save(file_name)

        with open(file_name, "rb") as f:
            st.download_button(
                label="📥 Tải xuống file Excel",
                data=f,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
