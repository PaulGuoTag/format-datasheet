import streamlit as st
from openpyxl import load_workbook
import io

def process_excel(file):
    # 加载工作簿
    wb = load_workbook(file)
    
    for ws in wb.worksheets:
        # 遍历所有有数据的单元格
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    # --- 步骤 A: 替换 [*] 为 / ---
                    # VBA 里的 " [*] " 左右有空格，这里完全照搬逻辑
                    val = cell.value.replace(" [*] ", "/")
                    
                    # --- 步骤 B: 清理开头的空格和斜杠 ---
                    # LTrim(cellVal) 后检查第一个字符是否为 "/"
                    stripped_val = val.lstrip()
                    if stripped_val.startswith("/"):
                        # 去掉开头的那个斜杠
                        val = stripped_val[1:]
                    
                    cell.value = val
    
    # 保存到内存流
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- Streamlit 界面 ---
st.set_page_config(page_title="Excel 批量清理工具")
st.title("🚀 Excel 数据清洗助手")
st.info("功能：将 ' [*] ' 替换为 '/'，并自动删除单元格开头的空格与斜杠。")

uploaded_files = st.file_uploader("请上传 Excel 文件 (支持多个)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        with st.spinner(f"正在处理 {uploaded_file.name}..."):
            processed_data = process_excel(uploaded_file)
            
            st.download_button(
                label=f"💾 下载已处理的 {uploaded_file.name}",
                data=processed_data,
                file_name=f"processed_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.success("所有文件处理完毕！")