import io
import importlib
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import streamlit as st

# Tự động tải lại các module nội bộ để tránh lỗi stale cache khi cập nhật code
for mod in [
    "src.config",
    "src.services.db_service",
    "src.services.memory_service",
    "src.services.gemini_service",
    "src.core.parser",
    "src.core.docx_translator",
    "src.core.xlsx_translator",
    "src.core.pdf_translator",
    "src.core.translator",
    "src.ui.styles",
]:
    if mod in sys.modules:
        try:
            importlib.reload(sys.modules[mod])
        except Exception:
            pass

from src.config import LANGS, AVAILABLE_MODELS, SAMPLE_PPTX, SUPPORTED_FORMATS
from src.core.translator import translate_document
from src.services.gemini_service import GeminiTranslator, load_glossary, ping_test_models
from src.services.memory_service import get_cache_stats, clear_cache, export_memory_markdown
from src.services.db_service import save_translated_task, get_recent_tasks, clear_all_tasks
from src.ui.styles import SINGLE_PAGE_CSS

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Document Translator AI • PPTX • Word • Excel • PDF",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Nhúng CSS thích ứng Sáng/Tối, nút lưu bên trên, fixed scroll log
st.markdown(SINGLE_PAGE_CSS, unsafe_allow_html=True)

# Khởi tạo session state
if "translated_data" not in st.session_state:
    st.session_state["translated_data"] = None
if "translated_filename" not in st.session_state:
    st.session_state["translated_filename"] = ""
if "report_text" not in st.session_state:
    st.session_state["report_text"] = ""
if "live_logs" not in st.session_state:
    st.session_state["live_logs"] = []
if "ping_results" not in st.session_state:
    st.session_state["ping_results"] = None

# Tiêu đề thanh lịch 1 dòng (Chữ sáng rực rỡ, độ tương phản cao)
st.markdown(
    '<div class="main-title">✨ <span class="title-text">AI Document Translator</span> '
    '<span class="badge-tag">v3.0 Multi-Format</span> '
    '<span class="main-sub">'
    'PPTX • Word • Excel • PDF — Bảo toàn 100% Layout, Bảng biểu, Công thức & Màu sắc'
    '</span></div>',
    unsafe_allow_html=True
)

# ----------------- THANH BÊN (SIDEBAR) -----------------
with st.sidebar:
    # 🌟 GHI CHÚ MODEL ĐANG DÙNG (CHỈ HIỆN KHI ĐÃ TEST API THÀNH CÔNG HOẶC ĐANG DỊCH)
    is_online = st.session_state.get("is_llm_online", False)
    current_active = st.session_state.get("active_model", "gemini-3.5-flash")
    if is_online:
        st.markdown(
            f"""
            <div class="active-model-card">
                <span class="active-model-name">🤖 {current_active}</span>
                <span class="active-model-badge">ONLINE 🟢</span>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### 🛠️ Cấu hình & Công cụ")
    default_key = os.getenv("GEMINI_API_KEY", "")
    if not default_key:
        try:
            if hasattr(st, "secrets") and "GEMINI_API_KEY" in st.secrets:
                default_key = st.secrets["GEMINI_API_KEY"]
        except Exception:
            pass

    api_key = st.text_input(
        "🔑 Gemini API Key", 
        value=default_key, 
        type="password",
        help="Tự động nạp từ file .env (máy tính) hoặc Streamlit Secrets (khi deploy Cloud)"
    )

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        btn_ping = st.button("🧪 Test API & Quota", use_container_width=True, help="Kiểm tra hạn mức Quota và thời gian reset")
    with col_btn2:
        btn_clear = st.button("🔄 Reset UI", use_container_width=True, help="Làm mới giao diện")

    if btn_clear:
        st.session_state["live_logs"] = []
        st.session_state["translated_data"] = None
        st.session_state["report_text"] = ""
        st.session_state["ping_results"] = None
        st.session_state["is_llm_online"] = False
        st.rerun()

    if btn_ping:
        if not api_key or not api_key.strip():
            st.error("❌ Chưa nhập API key!")
        else:
            with st.spinner("📡 Đang kiểm tra quota & thời gian reset..."):
                ping_res = ping_test_models(api_key.strip())
                st.session_state["ping_results"] = ping_res
                # Bật trạng thái Online nếu API key hợp lệ
                if any(r.get("status") in ("OK", "QUOTA_EXCEEDED") for r in ping_res):
                    st.session_state["is_llm_online"] = True
            st.rerun()

    # HIỂN THỊ CHI TIẾT QUOTA VÀ THỜI GIAN RESET
    if st.session_state.get("ping_results"):
        # Tự động dọn dẹp dữ liệu cũ của phiên trước nếu thiếu trường
        if not isinstance(st.session_state["ping_results"], list) or any(
            not isinstance(r, dict) or "quota" not in r for r in st.session_state["ping_results"]
        ):
            st.session_state["ping_results"] = None
        else:
            st.markdown("#### 📡 Trạng Thái Quota & Reset:")
            for r in st.session_state["ping_results"]:
                model_name = r.get("model", "Gemini Model")
                ping_val = r.get("ping", 0)
                usage_txt = r.get("usage_text", f"Đã dùng: {r.get('used', 0)}/{r.get('limit', 1500)} • Còn lại: {r.get('remaining', 1500)}")
                state_val = r.get("state", "🟢 Sẵn sàng")
                reset_val = r.get("reset", "Tự động làm mới hàng ngày")
                card_cls = r.get("card_class", "quota-card-ok")
                badge_cls = r.get("badge_class", "ready")
                pct_val = r.get("pct", 0)
                fill_cls = "fill-warn" if r.get("status") == "QUOTA_EXCEEDED" else "fill-ok"
                active_tag = '<span class="quota-active-tag">⭐ ĐANG DÙNG</span>' if model_name == current_active else ''
                
                card_html = f"""
                <div class="quota-card {card_cls}">
                    <div class="quota-header-row">
                        <span class="quota-model-name">🤖 {model_name}{active_tag}</span>
                        <span class="quota-ping-pill">⚡ {ping_val}ms</span>
                    </div>
                    <div class="quota-detail-line">📊 {usage_txt}</div>
                    <div class="quota-usage-track">
                        <div class="quota-usage-fill {fill_cls}" style="width: {pct_val}%;"></div>
                    </div>
                    <div class="quota-detail-line">🚦 <b>Trạng thái:</b> {state_val}</div>
                    <div class="quota-reset-badge {badge_cls}">
                        ⏰ {reset_val}
                    </div>
                </div>
                """
                st.markdown(card_html, unsafe_allow_html=True)

    model_choice = st.selectbox("🤖 Model LLM Ưu Tiên", AVAILABLE_MODELS, index=0)
    model = st.text_input("Tên Model", "gemini-3.5-flash") if model_choice == "Tùy chỉnh khác..." else model_choice
    if st.session_state.get("selected_model") != model:
        st.session_state["selected_model"] = model
        st.session_state["active_model"] = model

    st.markdown("---")
    st.markdown("### 📚 Bộ nhớ Song ngữ (TM)")
    use_cache = st.checkbox("Bật Bộ nhớ từ điển (SQLite TM)", value=True, help="Nếu tắt: Dịch thuần trong RAM, không lưu cache vào database.")
    batch_size = st.slider("Gói dịch (Đoạn/lô)", 20, 80, 50, 10, help="Gói lớn giúp tiết kiệm số lượt gọi API.")
    rate_delay = st.slider("Nghỉ an toàn (giây)", 0.0, 2.0, 0.3, 0.1)

    mem_stats = get_cache_stats()
    if mem_stats["total"] > 0:
        st.caption(f"⚡ `SQLite TM`: {mem_stats['total']} câu • {mem_stats['db_size_kb']} KB DB")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "📖 Tải .MD",
                data=export_memory_markdown(),
                file_name="translation_memory.md",
                mime="text/markdown",
                use_container_width=True,
                help="Xuất từ điển song ngữ trực tiếp từ SQLite ra file Markdown"
            )
        with c2:
            if st.button("🗑️ Xóa TM", use_container_width=True, help="Xóa sạch các câu đã lưu trong SQLite TM"):
                clear_cache()
                st.rerun()
    else:
        st.caption("⚡ `SQLite TM`: Đang trống (Sẽ tự động tích lũy khi bạn dịch file)")

    st.markdown("---")
    st.markdown("### 📐 Bố cục & Định dạng")
    st.caption("✅ **AI Smart Auto-Layout**: Tự động co chữ chống tràn khung cho PPTX, PDF & Excel.")
    with st.expander("Tùy chỉnh co chữ & Header"):
        fit = st.checkbox("Tự co chữ chống tràn", value=True)
        min_scale = st.slider("Co tối đa (%)", 50, 90, 65, 5) / 100
        layouts = st.checkbox("Dịch Master / Header & Footer", value=True)

# ----------------- KHUNG HÌNH CHÍNH (SINGLE-SCREEN 2 CỘT) -----------------
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown('<div class="section-title">📂 1. Chọn tệp & Thiết lập</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        "Kéo thả tệp PowerPoint, Word, Excel hoặc PDF (.pptx, .docx, .xlsx, .pdf)",
        type=["pptx", "docx", "xlsx", "pdf"],
        label_visibility="collapsed"
    )
    
    if not uploaded_file and SAMPLE_PPTX.exists():
        if st.button("💡 Dùng thử nhanh file mẫu có sẵn"):
            with open(SAMPLE_PPTX, "rb") as sf:
                uploaded_file = io.BytesIO(sf.read())
                uploaded_file.name = "sample_presentation.pptx"
                st.session_state["use_sample"] = uploaded_file

    if "use_sample" in st.session_state and not uploaded_file:
        uploaded_file = st.session_state["use_sample"]

    file_fmt_label, file_icon, _ = ("Tài liệu", "📄", "application/octet-stream")
    if uploaded_file:
        up_ext = Path(uploaded_file.name).suffix.lower()
        file_fmt_label, file_icon, _ = SUPPORTED_FORMATS.get(up_ext, ("Tài liệu", "📄", "application/octet-stream"))
        st.success(f"{file_icon} **{uploaded_file.name}** ({len(uploaded_file.getvalue())/1024:.1f} KB • {file_fmt_label})")

    # 🌐 CHỌN NGÔN NGỮ GỐC VÀ NGÔN NGỮ ĐÍCH
    col_l1, col_l2 = st.columns(2, gap="small")
    with col_l1:
        source_options = ["auto"] + list(LANGS.keys())
        source_lang = st.selectbox(
            "🔤 Ngôn ngữ gốc (Source):",
            options=source_options,
            format_func=lambda k: "🌐 Tự động nhận diện" if k == "auto" else f"{LANGS[k][0]} ({k})",
            index=0,
            help="Chọn ngôn ngữ gốc của tài liệu hoặc để AI tự động nhận diện"
        )
    with col_l2:
        default_target_idx = 0 if source_lang != "vi" else 1
        target_lang = st.selectbox(
            "🎯 Ngôn ngữ dịch sang (Target):",
            options=list(LANGS.keys()),
            format_func=lambda k: f"{LANGS[k][0]} ({k})",
            index=default_target_idx,
            help="Chọn ngôn ngữ bạn muốn dịch tài liệu sang"
        )

    src_label = "Tự động nhận diện" if source_lang == "auto" else LANGS[source_lang][0]
    dst_label = LANGS[target_lang][0]
    st.markdown(
        f'<div style="font-size: 0.8rem; color: var(--text-color); opacity: 0.82; margin: -4px 0 8px 0;">'
        f'🧭 Hướng dịch: <b>{src_label}</b> ➔ <b>{dst_label}</b>'
        f'</div>',
        unsafe_allow_html=True
    )
    if source_lang == target_lang and source_lang != "auto":
        st.warning("⚠️ Ngôn ngữ gốc và ngôn ngữ đích đang trùng nhau!")

    with st.expander("📖 Bảng thuật ngữ tùy biến (Glossary)"):
        gl_text = st.text_area("Mỗi dòng: `thuật ngữ = bản dịch`", placeholder="API = Giao diện lập trình", height=60, label_visibility="collapsed")

    btn_disabled = not (uploaded_file and api_key)
    btn_label = f"🚀 BẮT ĐẦU DỊCH {file_fmt_label.upper()}" if uploaded_file else "🚀 BẮT ĐẦU DỊCH TÀI LIỆU"
    btn_translate = st.button(btn_label, type="primary", disabled=btn_disabled, use_container_width=True)

    # 🔥 NHẬT KÝ THỜI GIAN THỰC (LIVE TERMINAL) NẰM NGAY DƯỚI NÚT DỊCH Ở CỘT TRÁI
    st.markdown('<div style="margin-top: 0.4rem;"></div>', unsafe_allow_html=True)
    log_box = st.empty()

    def render_logs(new_msg=None):
        if new_msg:
            timestamp = datetime.now().strftime("%H:%M:%S")
            st.session_state["live_logs"].append(f"[{timestamp}] {new_msg}")
        
        recent_logs = st.session_state["live_logs"][-40:]
        if not recent_logs:
            recent_logs = ["Chờ lệnh bắt đầu dịch từ bạn..."]
            
        lines_html = []
        for l in recent_logs:
            if any(k in l for k in ["✅", "Hoàn tất", "thành công", "100%"]):
                color = "#4ade80"  # Xanh lá
            elif "⚠️" in l:
                color = "#fbbf24"  # Vàng cam
            elif any(k in l for k in ["❌", "LỖI", "Error"]):
                color = "#f87171"  # Đỏ
            elif any(k in l for k in ["🚀", "Gói", "Bắt đầu"]):
                color = "#38bdf8"  # Xanh cyan
            elif any(k in l for k in ["🎨", "📐", "phục hồi", "bố cục"]):
                color = "#c084fc"  # Tím pastel
            else:
                color = "#cbd5e1"  # Xám bạc
            lines_html.append(f"<div style='color: {color}; margin-bottom: 2px;'>{l}</div>")
            
        terminal_html = f"""
        <div class="terminal-container">
            <div class="terminal-header">
                <div class="terminal-header-dots">
                    <span class="t-dot t-dot-red"></span>
                    <span class="t-dot t-dot-yellow"></span>
                    <span class="t-dot t-dot-green"></span>
                </div>
                <span class="terminal-header-title">LIVE TERMINAL LOG • REAL-TIME</span>
                <span class="terminal-header-badge">ONLINE 🟢</span>
            </div>
            <div class="terminal-box" id="term_log_box">
                {''.join(lines_html)}
            </div>
        </div>
        <script>
            var tb = document.getElementById("term_log_box");
            if (tb) {{ tb.scrollTop = tb.scrollHeight; }}
        </script>
        """
        log_box.markdown(terminal_html, unsafe_allow_html=True)

    render_logs()

with col_right:
    # 🔥 CỘT PHẢI: TIẾN ĐỘ & DANH SÁCH DATA CÁC FILE ĐÃ DỊCH (SQLITE DB)
    hdr_c1, hdr_c2 = st.columns([1.6, 1], gap="small")
    with hdr_c1:
        st.markdown('<div class="section-title">📊 2. TIẾN ĐỘ & FILE ĐÃ DỊCH</div>', unsafe_allow_html=True)
    with hdr_c2:
        if st.button("🗑️ Dọn lịch sử DB", help="Xóa danh sách file đã dịch trong SQLite DB", use_container_width=True):
            clear_all_tasks()
            st.session_state["translated_data"] = None
            st.rerun()
    
    progress_bar = st.progress(0.0)
    status_label = st.empty()

    # DANH SÁCH CÁC FILE ĐÃ DỊCH TRONG SQLITE DB KÈM NÚT TẢI RIÊNG CHO TỪNG FILE
    recent_tasks = get_recent_tasks(limit=30)
    if not recent_tasks:
        st.markdown(
            '<div class="empty-tasks-card">'
            '<div style="font-size: 1.8rem; margin-bottom: 6px;">📂</div>'
            '<div style="font-weight: 700; font-size: 0.95rem; margin-bottom: 4px;">Chưa có tệp nào được dịch</div>'
            '<div style="font-size: 0.8rem; opacity: 0.75;">Các file dịch hoàn thành sẽ tự động lưu vào SQLite DB và hiển thị tại đây kèm nút tải về riêng biệt.</div>'
            '</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f'<div style="font-size: 0.82rem; font-weight: 700; color: var(--text-color); margin: 0.2rem 0 0.4rem 0;">'
            f'📁 ĐÃ LƯU TRỮ {len(recent_tasks)} TỆP TRONG CƠ SỞ DỮ LIỆU SQLITE:'
            f'</div>',
            unsafe_allow_html=True
        )
        task_scroll_container = st.container(height=350)
        with task_scroll_container:
            for idx, t in enumerate(recent_tasks):
                tc1, tc2 = st.columns([3.2, 0.9], gap="small")
                t_ext = Path(t["translated_file"]).suffix.lower()
                _, t_icon, t_mime = SUPPORTED_FORMATS.get(t_ext, ("Tài liệu", "📄", "application/octet-stream"))
                with tc1:
                    lang_label = LANGS.get(t["target_lang"], (t["target_lang"],))[0]
                    new_tag = '<span class="new-task-pill">MỚI</span>' if idx == 0 else ''
                    st.markdown(
                        f"<div class='task-item-card'>"
                        f"{t_icon} <span class='task-file-title'>{t['translated_file']}</span> "
                        f"<span class='task-size-tag'>{t['file_size_kb']} KB</span>{new_tag}<br/>"
                        f"<span style='font-size: 0.70rem; opacity: 0.72;'>"
                        f"🌐 {lang_label} • ⏰ {t['created_at']}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with tc2:
                    if t["file_exists"]:
                        with open(t["file_path"], "rb") as tf:
                            st.download_button(
                                label="📥 Tải",
                                data=tf.read(),
                                file_name=t["translated_file"],
                                mime=t_mime,
                                key=f"dl_task_btn_{t['task_id']}_{idx}",
                                use_container_width=True
                            )
                    else:
                        st.caption("Đã xóa")
                st.markdown("<hr style='margin: 2px 0 4px 0; opacity: 0.08;'/>", unsafe_allow_html=True)

    if btn_translate:
        st.session_state["live_logs"] = []
        st.session_state["translated_data"] = None
        st.session_state["is_llm_online"] = True
        render_logs(f"🚀 Bắt đầu tiếp nhận tệp: {uploaded_file.name}")
        
        def ui_progress(p, msg):
            progress_bar.progress(p)
            status_label.markdown(f"⏳ **Tiến độ {int(p * 100)}%:** {msg}")

        def ui_log(msg):
            render_logs(msg)

        try:
            with tempfile.TemporaryDirectory() as d:
                up_ext = Path(uploaded_file.name).suffix.lower() or ".pptx"
                src_path = os.path.join(d, f"input{up_ext}")
                dst_path = os.path.join(d, f"translated{up_ext}")
                
                with open(src_path, "wb") as f_out:
                    f_out.write(uploaded_file.getvalue())
                
                ui_log(f"🤖 Khởi tạo bộ dịch Gemini AI: {model}")
                translator = GeminiTranslator(api_key=api_key, model=model, log_func=ui_log)
                glossary = load_glossary(gl_text) if gl_text.strip() else None
                
                report = translate_document(
                    src=src_path,
                    dst=dst_path,
                    target=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    glossary=glossary,
                    fit=fit,
                    min_scale=min_scale,
                    include_layouts=layouts,
                    progress=ui_progress,
                    log_func=ui_log,
                    use_cache=use_cache,
                    batch_size=batch_size,
                    rate_delay=rate_delay
                )
                
                # Cập nhật model thực tế đã xử lý hoàn tất
                st.session_state["active_model"] = getattr(translator, "active_model", model)
                
                with open(dst_path, "rb") as f_res:
                    st.session_state["translated_data"] = f_res.read()
                
                base_name = Path(uploaded_file.name).stem
                st.session_state["translated_filename"] = f"{base_name}.{target_lang}{up_ext}"
                st.session_state["report_text"] = report.text()
                
                # Lưu vào SQLite Database & thư mục exports
                save_translated_task(
                    source_name=uploaded_file.name,
                    translated_name=st.session_state["translated_filename"],
                    target_lang=target_lang,
                    slides_count=getattr(report, "units", 0),
                    units_count=getattr(report, "unique", 0),
                    file_bytes=st.session_state["translated_data"]
                )
                
            progress_bar.progress(1.0)
            status_label.markdown(f"🎉 **Hoàn thành 100%! Đã thêm '{st.session_state['translated_filename']}' vào danh sách file.**")
            ui_log("🎉 Dịch thành công! File đã được lưu vào SQLite DB và danh sách bên phải.")
            st.rerun()
        except Exception as e:
            ui_log(f"❌ LỖI: {e}")
            st.error(f"❌ Có lỗi xảy ra: {e}")
