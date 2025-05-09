import streamlit as st
import pandas as pd
import os
import shutil
from io import BytesIO
from zipfile import ZipFile

st.set_page_config(page_title="Excel 渠道拆分工具", layout="wide")
st.title("📊 Excel 渠道拆分工具")

uploaded_file = st.file_uploader("请上传 Excel 文件", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.success("文件上传成功，以下为预览：")
    st.dataframe(df.head())

    columns = df.columns.tolist()
    col1, col2, col3, col4 = st.columns(4)
    渠道列 = col1.selectbox("选择 渠道列", columns)
    会议列 = col2.selectbox("选择 会议信息列", columns)
    paper_id列 = col3.selectbox("选择 paper_id 列", columns)
    volume列 = col4.selectbox("选择 volume 列", columns)
    link列 = st.selectbox("选择 见刊链接列", columns)

    阈值 = st.number_input("最小保留行数（小于此数将文本导出）", min_value=1, value=5)

    所有渠道 = df[渠道列].dropna().unique().tolist()
    选中渠道 = st.multiselect("选择要处理的渠道", 所有渠道, default=所有渠道)

    if st.button("🚀 立即拆分生成"):
        output_dir = "output"
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir, exist_ok=True)

        text_output = ""

        for 渠道 in 选中渠道:
            df_渠道 = df[df[渠道列] == 渠道]
            if len(df_渠道) < 阈值:
                text_output += f"{渠道}\n"
                for _, row in df_渠道.iterrows():
                    pid = row[paper_id列]
                    vol = row[volume列]
                    link = row[link列]
                    text_output += f"{pid}   {vol}\n{link}\n\n"
            else:
                filename = os.path.join(output_dir, f"{渠道}.xlsx")
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df_渠道.to_excel(writer, sheet_name='汇总', index=False)
                    df_渠道['主会议'] = df_渠道[会议列].apply(lambda x: str(x).split('_')[0])
                    for 会名, df_会 in df_渠道.groupby(会议列):
                        df_会.to_excel(writer, sheet_name=str(会名)[:31], index=False)

        if text_output:
            with open(os.path.join(output_dir, "少于阈值渠道信息.txt"), "w", encoding='utf-8') as f:
                f.write(text_output)

        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zip_file:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    filepath = os.path.join(root, file)
                    zip_file.write(filepath, arcname=os.path.relpath(filepath, output_dir))

        st.success("✅ 拆分完成！点击下载结果：")
        st.download_button("📥 下载 ZIP 文件", zip_buffer.getvalue(), file_name="渠道拆分结果.zip")
