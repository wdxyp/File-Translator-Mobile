import os
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

import baidu_Translator_v4_0 as bt


st.set_page_config(page_title="百度长语句翻译", layout="centered")

st.title("百度长语句翻译")

col1, col2 = st.columns(2)
with col1:
    app_id = st.text_input("BAIDU_APP_ID", value=os.getenv("BAIDU_APP_ID", ""))
with col2:
    secret_key = st.text_input("BAIDU_SECRET_KEY", value=os.getenv("BAIDU_SECRET_KEY", ""), type="password")

direction = st.selectbox("翻译方向", options=sorted(bt.TO_LANG_MAP.keys()), index=sorted(bt.TO_LANG_MAP.keys()).index("ko2zh") if "ko2zh" in bt.TO_LANG_MAP else 0)
append_translation = st.checkbox("在原文下方保留翻译对照", value=False)
generate_corpus = st.checkbox("生成语料库（Corpus）", value=False)

uploaded_file = st.file_uploader("上传文件（docx / xlsx / pptx / ppt）", type=["docx", "xlsx", "pptx", "ppt"])
uploaded_revision = st.file_uploader("可选：上传 revision.md（校准词典）", type=["md", "txt"])

run = st.button("开始翻译", type="primary", disabled=(uploaded_file is None))

if run and uploaded_file is not None:
    with st.spinner("翻译中…"):
        try:
            with tempfile.TemporaryDirectory() as work_dir:
                work_dir_path = Path(work_dir)
                in_path = work_dir_path / uploaded_file.name
                in_path.write_bytes(uploaded_file.getvalue())

                out_dir = work_dir_path / "out"
                out_dir.mkdir(parents=True, exist_ok=True)

                default_revision_path = Path(__file__).with_name("revision.md")
                if uploaded_revision is not None:
                    revision_path = work_dir_path / "revision.md"
                    revision_path.write_bytes(uploaded_revision.getvalue())
                else:
                    revision_path = default_revision_path

                prev_cwd = os.getcwd()
                os.chdir(work_dir)
                try:
                    output_path = bt.translate_file(
                        input_file=str(in_path),
                        output_dir=str(out_dir),
                        name=f"translated_{Path(uploaded_file.name).stem}",
                        direction=direction,
                        append=append_translation,
                        corpus=generate_corpus,
                        revision_file=str(revision_path),
                        app_id=app_id or None,
                        secret_key=secret_key or None,
                    )
                finally:
                    os.chdir(prev_cwd)

                output_file_path = Path(output_path)
                output_bytes = output_file_path.read_bytes()

                if generate_corpus:
                    corpus_dir = work_dir_path / "Corpus"
                    zip_path = work_dir_path / "result.zip"
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr(output_file_path.name, output_bytes)
                        if corpus_dir.exists():
                            for p in sorted(corpus_dir.glob("*.xlsx")):
                                zf.write(p, arcname=f"Corpus/{p.name}")
                    st.download_button(
                        "下载结果（ZIP）",
                        data=zip_path.read_bytes(),
                        file_name="result.zip",
                        mime="application/zip",
                    )
                else:
                    st.download_button(
                        "下载翻译后的文件",
                        data=output_bytes,
                        file_name=output_file_path.name,
                        mime="application/octet-stream",
                    )

                st.success("完成")
        except Exception as e:
            st.error(str(e))
