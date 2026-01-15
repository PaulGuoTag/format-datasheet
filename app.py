import streamlit as st
from openpyxl import load_workbook
import io
import re
import zipfile

# 页面配置
st.set_page_config(page_title="Excel 强力清洗工具", layout="centered")

def process_excel(file_content):
    wb = load_workbook(io.BytesIO(file_content))
    
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val = cell.value
                    
                    # 1. 预处理特殊空格 (\xa0 -> 普通空格)
                    val = val.replace('\xa0', ' ')
                    
                    # -------------------------------------------------------
                    # 2. 核心修正：使用正则通配符匹配 [任意内容]
                    # -------------------------------------------------------
                    # r"\s*"  -> 匹配左右可能存在的空格
                    # r"\["   -> 匹配左中括号
                    # r".*?"  -> 匹配中间的任意字符 (数字、字母等)
                    # r"\]"   -> 匹配右中括号
                    val = re.sub(r"\s*\[.*?\]\s*", "/", val)
                    
                    # 3. 循环清理开头 (去除开头的空格和斜杠)
                    while True:
                        temp = val.lstrip() 
                        if temp.startswith("/"):
                            val = temp[1:] 
                        else:
                            val = temp
                            break
                    
                    cell.value = val
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 界面部分 ---
st.title("🚀 Excel 数据清洗 (支持通配符)")
st.info("当前逻辑：匹配 `[任何内容]` (如 `[001]`, `[AB-9]`) 并替换为 `/`，同时清理开头。")

uploaded_files = st.file_uploader("上传文件", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    processed_files = {} 
    
    progress_bar = st.progress(0)
    for index, uploaded_file in enumerate(uploaded_files):
        with st.spinner(f"正在清洗: {uploaded_file.name}"):
            file_bytes = uploaded_file.read()
            output_data = process_excel(file_bytes)
            processed_files[f"processed_{uploaded_file.name}"] = output_data
        progress_bar.progress((index + 1) / len(uploaded_files))

    st.success("处理完成！")

    if len(processed_files) == 1:
        file_name, data = list(processed_files.items())[0]
        st.download_button("💾 下载结果", data, file_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
            for name, data in processed_files.items():
                zf.writestr(name, data)
        st.download_button("📦 下载 ZIP 包", zip_buffer.getvalue(), "processed_files.zip", "application/zip", type="primary")