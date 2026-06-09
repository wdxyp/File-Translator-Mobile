import os
import re
import json
import tempfile
import time
import uuid
import zipfile
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from threading import Lock

import requests
import streamlit as st

import baidu_translator_v212 as baidu_bt
import kimi_translator_v214 as kimi_bt


APP_TITLE = "My File Trans"
APP_DIR = Path(__file__).resolve().parent
GITHUB_REVISION_URL = "https://raw.githubusercontent.com/wdxyp/File-Translator-Mobile/main/revision.md"
APP_VERSION = os.getenv("APP_VERSION", "").strip()
APP_VERSION_JSON = ""
try:
    version_path = APP_DIR / "version.json"
    if version_path.exists():
        data = json.loads(version_path.read_text(encoding="utf-8"))
        APP_VERSION_JSON = str(data.get("version", "")).strip()
except Exception:
    APP_VERSION_JSON = ""

try:
    os.chdir(APP_DIR)
except Exception:
    pass

st.set_page_config(page_title=APP_TITLE, layout="centered")

if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
if "result_file_name" not in st.session_state:
    st.session_state.result_file_name = None
if "result_mime" not in st.session_state:
    st.session_state.result_mime = None
if "log_lines" not in st.session_state:
    st.session_state.log_lines = []
if "translate_requested" not in st.session_state:
    st.session_state.translate_requested = False
if "is_running" not in st.session_state:
    st.session_state.is_running = False
if "session_id" not in st.session_state:
    st.session_state.session_id = uuid.uuid4().hex
if "lock_acquired" not in st.session_state:
    st.session_state.lock_acquired = False
if "show_busy_dialog" not in st.session_state:
    st.session_state.show_busy_dialog = False
if "progress_pct" not in st.session_state:
    st.session_state.progress_pct = 0
if "progress_text" not in st.session_state:
    st.session_state.progress_text = ""
if "download_clicked" not in st.session_state:
    st.session_state.download_clicked = False
if "uploader_nonce" not in st.session_state:
    st.session_state.uploader_nonce = 0

if st.session_state.is_running and (not st.session_state.translate_requested or not st.session_state.lock_acquired):
    st.session_state.is_running = False

@st.cache_resource
def _global_translate_lock():
    return {"lock": Lock(), "owner_session_id": None, "started_at": None}

def _busy_dialog():
    if not st.session_state.show_busy_dialog:
        return
    try:
        dialog = st.dialog("正在使用中")
    except Exception:
        st.warning("正在使用中：已有其他用户在翻译，请稍后再试。")
        st.session_state.show_busy_dialog = False
        return

    @dialog
    def _render():
        st.write("已有其他用户正在翻译，请稍后再试。")
        if st.button("知道了", key="busy_dialog_ok", use_container_width=True):
            st.session_state.show_busy_dialog = False

    _render()

def _is_mobile():
    try:
        ctx = getattr(st, "context", None)
        headers = getattr(ctx, "headers", None) if ctx is not None else None
        ua = headers.get("User-Agent", "") if headers else ""
    except Exception:
        ua = ""
    return bool(re.search(r"(Mobile|Android|iPhone|iPad|iPod)", ua, flags=re.IGNORECASE))

st.title(APP_TITLE)
display_version = APP_VERSION_JSON or APP_VERSION or "2.0.0"
if display_version:
    if not display_version.lower().startswith("v"):
        display_version = f"v{display_version}"
    st.caption(f"Version: {display_version}")

# --- 引擎选择 ---
engine = st.radio("选择翻译引擎", ["Baidu (V2.12)", "Kimi (V2.14)"], horizontal=True, key="engine")

# --- 密钥配置 ---
if engine == "Baidu (V2.12)":
    app_id = os.getenv("BAIDU_APP_ID", "").strip()
    secret_key = os.getenv("BAIDU_SECRET_KEY", "").strip()
    
    if not app_id or not secret_key:
        st.error("未检测到 BAIDU_APP_ID / BAIDU_SECRET_KEY。请在 Streamlit Secrets 中配置。")
    
    # 语言映射使用 Baidu 的
    TO_LANG_MAP = baidu_bt.TO_LANG_MAP
else:
    kimi_api_key = os.getenv("KIMI_API_KEY", "").strip()
    kimi_model = os.getenv("KIMI_MODEL", "kimi-k2.6").strip()
    
    if not kimi_api_key:
        st.error("未检测到 KIMI_API_KEY。请在 Streamlit Secrets 中配置。")
        
    # 语言映射使用 Kimi 的（通常两者一致，但以对应引擎为准）
    TO_LANG_MAP = kimi_bt.TO_LANG_MAP

direction = st.selectbox(
    "翻译方向",
    options=sorted(TO_LANG_MAP.keys()),
    index=sorted(TO_LANG_MAP.keys()).index("ko2zh") if "ko2zh" in TO_LANG_MAP else 0,
    key="direction",
)
append_translation = st.checkbox("在原文下方保留翻译对照", value=False, key="append_translation")
generate_corpus = st.checkbox("生成语料库（Corpus）", value=False, key="generate_corpus")

uploaded_file = st.file_uploader(
    "上传文件（docx / xlsx / pptx / ppt）",
    type=["docx", "xlsx", "pptx", "ppt"],
    key=f"uploaded_file_{st.session_state.uploader_nonce}",
)

use_github_revision = st.checkbox("使用 GitHub 上的 revision.md（自动更新）", value=True, key="use_github_revision")

# --- 翻译方向预检 ---
direction_warning = None
if uploaded_file is not None:
    # 使用缓存避免重复解析文件内容，提高切换方向时的响应速度
    @st.cache_data(show_spinner=False)
    def _get_file_languages(file_bytes, file_name):
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file_name).suffix) as tmp:
                tmp.write(file_bytes)
                tmp_path = tmp.name
            
            # 借用百度插件的逻辑来检测语言（两者逻辑一致）
            # 我们只需要检测一次语言特征
            sample_text = baidu_bt._extract_sample_text(tmp_path)
            if not sample_text:
                return None
            
            return {
                "has_korean": bool(baidu_bt.HANGUL_RE.search(sample_text)),
                "has_chinese": bool(baidu_bt.CHINESE_RE.search(sample_text)),
                "has_japanese": bool(re.search(r'[\u3040-\u30ff]', sample_text)),
                "has_english": bool(re.search(r'[a-zA-Z]{3,}', sample_text))
            }
        except Exception:
            return None
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try: os.remove(tmp_path)
                except Exception: pass

    st.info(f"正在预检文件内容语言: {uploaded_file.name}...")
    
    # 获取文件语言特征
    langs = _get_file_languages(uploaded_file.getvalue(), uploaded_file.name)
    
    if langs:
        # 借用当前引擎的判定逻辑
        if engine == "Baidu (V2.12)":
            direction_warning = baidu_bt._judge_direction_warning(langs, direction)
        else:
            direction_warning = kimi_bt._judge_direction_warning(langs, direction)

if direction_warning:
    st.error(f"⚠️ 预检拦截：{direction_warning}。请切换到正确的翻译方向后重试！")

is_mobile = _is_mobile()

def _on_download_complete():
    st.session_state.result_bytes = None
    st.session_state.result_file_name = None
    st.session_state.result_mime = None
    st.session_state.log_lines = []
    st.session_state.progress_pct = 0
    st.session_state.progress_text = ""
    st.session_state.translate_requested = False
    st.session_state.is_running = False
    st.session_state.download_clicked = False
    st.session_state.show_busy_dialog = False
    st.session_state.lock_acquired = False
    st.session_state.uploader_nonce += 1

def _on_start_translate():
    if st.session_state.is_running:
        return

    g = _global_translate_lock()
    acquired = g["lock"].acquire(blocking=False)
    if not acquired:
        st.session_state.show_busy_dialog = True
        return

    g["owner_session_id"] = st.session_state.session_id
    g["started_at"] = time.time()
    st.session_state.lock_acquired = True
    st.session_state.result_bytes = None
    st.session_state.result_file_name = None
    st.session_state.result_mime = None
    st.session_state.log_lines = []
    st.session_state.download_clicked = False
    st.session_state.progress_pct = 0
    st.session_state.progress_text = "翻译中…"
    st.session_state.translate_requested = True
    st.session_state.is_running = True

col_run, col_dl, col_done = st.columns([2, 1, 1] if is_mobile else [1, 1, 1])

_busy_dialog()

with col_run:
    if engine == "Baidu (V2.12)":
        run_disabled = (
            st.session_state.is_running
            or bool(st.session_state.result_bytes and st.session_state.result_file_name)
            or uploaded_file is None
            or (not app_id or not secret_key)
            or direction_warning is not None
        )
    else:
        run_disabled = (
            st.session_state.is_running
            or bool(st.session_state.result_bytes and st.session_state.result_file_name)
            or uploaded_file is None
            or (not kimi_api_key)
            or direction_warning is not None
        )

    st.button(
        "开始翻译",
        disabled=run_disabled,
        use_container_width=True,
        key="run_translate",
        on_click=_on_start_translate,
    )
with col_dl:
    download_slot = st.empty()
with col_done:
    download_done_slot = st.empty()

progress_slot = st.empty()
log_slot = st.empty()

def _render_progress():
    if st.session_state.is_running:
        progress_slot.progress(int(st.session_state.progress_pct))
        if st.session_state.progress_text:
            progress_slot.caption(st.session_state.progress_text)
    else:
        progress_slot.empty()

def _render_download():
    if not (st.session_state.result_bytes and st.session_state.result_file_name):
        download_slot.button(
            "下载结果",
            disabled=True,
            use_container_width=True,
            key="download_result_disabled",
        )
        return
    clicked = download_slot.download_button(
        "下载结果",
        data=st.session_state.result_bytes,
        file_name=st.session_state.result_file_name,
        mime=st.session_state.result_mime or "application/octet-stream",
        use_container_width=True,
        key="download_result_ready",
    )
    if clicked:
        st.session_state.download_clicked = True

def _render_download_done():
    result_ready = bool(st.session_state.result_bytes and st.session_state.result_file_name)
    download_done_slot.button(
        "下载完成",
        disabled=(st.session_state.is_running or (not result_ready) or (not st.session_state.download_clicked)),
        use_container_width=True,
        key="download_done",
        on_click=_on_download_complete,
    )

def _render_log():
    lines = st.session_state.log_lines[-200:]
    log_slot.code("\n".join(lines))

_render_download()
_render_download_done()
_render_progress()
_render_log()

if st.session_state.translate_requested and uploaded_file is not None and st.session_state.lock_acquired:
    class _StreamlitLogWriter:
        def __init__(self):
            self._buf = ""
            self._last_render = 0.0
            self._progress_re = re.compile(r"第\\s*(\\d+)\\s*/\\s*(\\d+)")

        def write(self, s):
            if not s:
                return 0
            self._buf += s
            while "\n" in self._buf:
                line, self._buf = self._buf.split("\n", 1)
                self._handle_line(line.rstrip("\r"))
            return len(s)

        def flush(self):
            if self._buf.strip():
                self._handle_line(self._buf.rstrip("\r"))
                self._buf = ""

        def _handle_line(self, line):
            text = (line or "").strip()
            if not text:
                return

            st.session_state.log_lines.append(text)
            st.session_state.log_lines = st.session_state.log_lines[-2000:]

            m = self._progress_re.search(text)
            if m:
                cur = int(m.group(1))
                total = int(m.group(2))
                if total > 0:
                    st.session_state.progress_pct = max(0, min(99, int(cur * 100 / total)))
                    st.session_state.progress_text = text

            now = time.monotonic()
            if now - self._last_render < 0.3:
                return
            self._last_render = now

            _render_progress()
            _render_log()

    with st.spinner("翻译中…"):
        g = _global_translate_lock()
        try:
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
                try:
                    os.chdir(work_dir)
                except FileNotFoundError:
                    os.chdir(APP_DIR)
                    os.chdir(work_dir)
                try:
                    writer = _StreamlitLogWriter()
                    with redirect_stdout(writer), redirect_stderr(writer):
                        if engine == "Baidu (V2.12)":
                            output_path = baidu_bt.translate_file(
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
                        else:
                            output_path = kimi_bt.translate_file(
                                input_file=str(in_path),
                                output_dir=str(out_dir),
                                name=f"translated_{Path(uploaded_file.name).stem}",
                                direction=direction,
                                append=append_translation,
                                corpus=generate_corpus,
                                revision_file=str(revision_path),
                                api_key=kimi_api_key or None,
                                model=kimi_model or None,
                            )
                    writer.flush()
                finally:
                    try:
                        os.chdir(prev_cwd)
                    except FileNotFoundError:
                        os.chdir(APP_DIR)

                output_file_path = Path(output_path)
                output_bytes = output_file_path.read_bytes()

                if generate_corpus:
                    corpus_dir = work_dir_path / "Corpus"
                    corpus_files = []
                    if corpus_dir.exists():
                        corpus_files.extend(sorted(corpus_dir.glob("*.xlsx")))
                    if not corpus_files:
                        corpus_files.extend(sorted(work_dir_path.rglob("Corpus_v*.xlsx")))
                        corpus_files.extend(sorted(out_dir.rglob("Corpus_v*.xlsx")))
                    zip_path = work_dir_path / "result.zip"
                    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr(output_file_path.name, output_bytes)
                        for p in sorted({str(p) for p in corpus_files}):
                            p = Path(p)
                            zf.write(p, arcname=f"Corpus/{p.name}")
                    st.session_state.result_bytes = zip_path.read_bytes()
                    st.session_state.result_file_name = "result.zip"
                    st.session_state.result_mime = "application/zip"
                else:
                    st.session_state.result_bytes = output_bytes
                    st.session_state.result_file_name = output_file_path.name
                    st.session_state.result_mime = "application/octet-stream"

                st.session_state.is_running = False
                st.session_state.translate_requested = False
                st.session_state.progress_pct = 100
                st.session_state.progress_text = "完成"
                st.session_state.log_lines.append("翻译完成！")
                st.success("完成")
                st.rerun()
        except Exception as e:
            st.session_state.is_running = False
            st.session_state.translate_requested = False
            st.error(str(e))
        finally:
            if st.session_state.lock_acquired:
                try:
                    if g.get("owner_session_id") == st.session_state.session_id:
                        g["owner_session_id"] = None
                        g["started_at"] = None
                    g["lock"].release()
                except Exception:
                    pass
                st.session_state.lock_acquired = False
