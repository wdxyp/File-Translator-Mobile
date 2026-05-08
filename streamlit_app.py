import os
import tempfile
import zipfile
from pathlib import Path

import requests
import streamlit as st

import baidu_Translator_v4_0 as bt


APP_TITLE = "My File Trans"
GITHUB_REVISION_URL = "https://raw.githubusercontent.com/wdxyp/File-Translator-Mobile/main/revision.md"

st.set_page_config(page_title=APP_TITLE, layout="centered")

if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
if "result_file_name" not in st.session_state:
    st.session_state.result_file_name = None
if "result_mime" not in st.session_state:
    st.session_state.result_mime = None

st.title(APP_TITLE)

app_id = os.getenv("BAIDU_APP_ID", "").strip()
secret_key = os.getenv("BAIDU_SECRET_KEY", "").strip()
if not app_id or not secret_key:
    st.error("未检测到 BAIDU_APP_ID / BAIDU_SECRET_KEY。请在 Streamlit Cloud 的 Settings → Secrets 中配置后刷新页面。")

direction = st.selectbox("翻译方向", options=sorted(bt.TO_LANG_MAP.keys()), index=sorted(bt.TO_LANG_MAP.keys()).index("ko2zh") if "ko2zh" in bt.TO_LANG_MAP else 0)
append_translation = st.checkbox("在原文下方保留翻译对照", value=False)
generate_corpus = st.checkbox("生成语料库（Corpus）", value=False)

uploaded_file = st.file_uploader("上传文件（docx / xlsx / pptx / ppt）", type=["docx", "xlsx", "pptx", "ppt"])

use_github_revision = st.checkbox("使用 GitHub 上的 revision.md（自动更新）", value=True)

run = st.button("开始翻译", type="primary", disabled=(uploaded_file is None or (not app_id or not secret_key)))

if st.session_state.result_bytes and st.session_state.result_file_name:
    st.download_button(
        "下载结果",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_file_name,
        mime=st.session_state.result_mime or "application/octet-stream",
    )

if run and uploaded_file is not None:
    with st.spinner("翻译中…"):
        try:
            st.session_state.result_bytes = None
            st.session_state.result_file_name = None
            st.session_state.result_mime = None

            with tempfile.TemporaryDirectory() as work_dir:
                work_dir_path = Path(work_dir)
                in_path = work_dir_path / uploaded_file.name
                in_path.write_bytes(uploaded_file.getvalue())

                out_dir = work_dir_path / "out"
                out_dir.mkdir(parents=True, exist_ok=True)

                default_revision_path = Path(__file__).with_name("revision.md")
                revision_path = default_revision_path
                if use_github_revision:
                    try:
                        resp = requests.get(GITHUB_REVISION_URL, timeout=10)
                        resp.raise_for_status()
                        revision_path = work_dir_path / "revision.md"
                        revision_path.write_bytes(resp.content)
                    except Exception:
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
                    st.session_state.result_bytes = zip_path.read_bytes()
                    st.session_state.result_file_name = "result.zip"
                    st.session_state.result_mime = "application/zip"
                else:
                    st.session_state.result_bytes = output_bytes
                    st.session_state.result_file_name = output_file_path.name
                    st.session_state.result_mime = "application/octet-stream"

                st.success("完成")
        except Exception as e:
            st.error(str(e))
