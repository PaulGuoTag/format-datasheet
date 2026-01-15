import streamlit as st
from openpyxl import load_workbook
import io
import re
import zipfile

# 设置页面标题
st.set_page_config(page_title="Excel 批量清理工具", layout="centered")

def process_excel(file_content):
    """处理单个 Excel 文件的逻辑"""
    # 将上传的文件流载入 openpyxl
    wb = load_workbook(io.BytesIO(file_content))
    
    for ws in wb.worksheets:
        # 遍历所有有数据的单元格
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    val = cell.value
                    
                    # --- 步骤 A: 替换 [*] 为 / (处理各种空格情况) ---
                    # 正则解释：\s* 匹配零个或多个空格；\[\*\] 匹配字面量 [*]
                    val = re.sub(r"\s*\[\*\]\s*", "/", val)
                    
                    # --- 步骤 B: 清理开头的空格和斜杠 ---
                    val = val.lstrip()  # 去掉左侧空格
                    if val.startswith("/"):
                        val = val[1:].lstrip()  # 去掉斜杠后再洗一遍开头的空格
                    
                    cell.value = val
    
    # 保存处理后的文件到内存
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 界面部分 ---
st.title("🚀 Excel 数据清洗助手 (Web版)")
st.markdown("""
**功能说明：**
1. 将所有 `[*]`, ` [*] `, `[* ]` 等变体统一替换为 `/`。
2. 自动剔除单元格内容开头的空格和斜杠（例如 `/ 数据` 变为 `数据`）。
""")

# 文件上传组件
uploaded_files = st.file_uploader("请上传 Excel 文件 (支持多选)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    processed_files = {} # 存储处理后的文件数据 {文件名: 数据}
    
    with st.status("正在处理文件...", expanded=True) as status:
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.read()
            processed_data = process_excel(file_bytes)
            processed_files[f"processed_{uploaded_file.name}"] = processed_data
            st.write(f"✅ {uploaded_file.name} 处理完成")
        status.update(label="处理完毕!", state="complete", expanded=False)

    # 如果只有一个文件，直接提供下载
    if len(processed_files) == 1:
        file_name, data = list(processed_files.items())[0]
        st.download_button(
            label="💾 下载处理后的文件",
            data=data,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    
    # 如果有多个文件，打包成 ZIP 下载
    elif len(processed_files) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zf:
            for name, data in processed_files.items():
                zf.writestr(name, data)
        
        st.download_button(
            label="📦 一键下载所有文件的 ZIP 包",
            data=zip_buffer.getvalue(),
            file_name="processed_files.zip",
            mime="application/zip",
            type="primary"
        )

st.divider()
st.caption("提示：本工具在内存中处理，不会保存您的原始文件，关闭网页后数据即刻消失。")