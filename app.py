import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

def load_excel(file):
    try:
        df = pd.read_excel(file, engine='openpyxl')
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
        return None

def generate_excel_report(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        summary = df.describe()
        summary.to_excel(writer, sheet_name='Summary')
    output.seek(0)
    return output

st.set_page_config(page_title="Báo cáo dữ liệu", layout="wide")
st.title("Ứng dụng Báo cáo Dữ liệu với Streamlit")

uploaded_file = st.file_uploader("Tải lên file Excel", type=["xlsx"])

if uploaded_file:
    df = load_excel(uploaded_file)
    if df is not None:
        st.subheader("Xem trước dữ liệu")
        st.dataframe(df.head())
        
        st.subheader("Thông tin dữ liệu")
        st.write(f"Số hàng: {df.shape[0]}, Số cột: {df.shape[1]}")
        st.write("Số lượng giá trị trống:")
        st.dataframe(df.isnull().sum())
        
        st.subheader("Thống kê mô tả")
        st.dataframe(df.describe())
        
        st.subheader("Trực quan hóa dữ liệu")
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            selected_col = st.selectbox("Chọn cột số để vẽ biểu đồ:", numeric_cols)
            fig, ax = plt.subplots()
            sns.histplot(df[selected_col], bins=20, kde=True, ax=ax)
            st.pyplot(fig)
        else:
            st.write("Không có dữ liệu số để vẽ biểu đồ.")
        
        excel_report = generate_excel_report(df)
        st.download_button(
            label="📥 Tải xuống báo cáo Excel",
            data=excel_report,
            file_name="report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
