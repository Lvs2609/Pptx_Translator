import io
import importlib
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path
import streamlit as st

# Tự động tải lại các module nội bộ để tránh lỗi stale cache khi cập nhật code
for mod in ["src.services.memory_service", "src.services.db_service", "src.services.gemini_service", "src.ui.styles"]:
    if mod in sys.modules:
        try:
            importlib.reload(sys.modules[mod])
        except Exception:
            pass

from src.config import LANGS, AVAILABLE_MODELS, SAMPLE_PPTX
from src.core.translator import translate_pptx
from src.services.gemini_service import GeminiTranslator, load_glossary, ping_test_models
from src.services.memory_service import get_cache_stats, clear_cache, export_memory_markdown
from src.services.db_service import save_translated_task, get_recent_tasks, clear_all_tasks
from src.ui.styles import SINGLE_PAGE_CSS

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="PPTX Translator AI",
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
    '<div class="main-title">✨ <span class="title-text">PPTX Translator AI</span> '
    '<span class="badge-tag">v2.5 Pro Engine</span> '
    '<span class="main-sub">'
    'Bảo toàn 100% Layout, Hình ảnh, Số liệu & Màu sắc chữ'
    '</span></div>',
    unsafe_allow_html=True
)

# ----------------- THANH BÊN (SIDEBAR) -----------------
with st.sidebar:
    st.markdown("### 🛠️ Cấu hình & Công cụ")
    api_key = st.text_input(
        "🔑 Gemini API Key", 
        value=os.getenv("GEMINI_API_KEY", ""), 
        type="password",
        help="Đã tự động nạp từ file .env"
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
        st.rerun()

    if btn_ping:
        if not api_key:
            st.error("❌ Chưa có API key!")
        else:
            with st.status("📡 Đang kiểm tra quota & thời gian reset..."):
                st.session_state["ping_results"] = ping_test_models(api_key)

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
                
                card_html = f"""
                <div class="quota-card {card_cls}">
                    <div class="quota-header-row">
                        <span class="quota-model-name">🤖 {model_name}</span>
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
    st.markdown("### 📐 Bố cục Slide")
    st.caption("✅ **AI Smart Auto-Layout**: Tự động co chữ vừa khít khung, chống tràn viền.")
    with st.expander("Tùy chỉnh co chữ"):
        fit = st.checkbox("Tự co chữ", value=True)
        min_scale = st.slider("Co tối đa (%)", 50, 90, 65, 5) / 100
        layouts = st.checkbox("Dịch cả slide master", value=True)

# ----------------- KHUNG HÌNH CHÍNH (SINGLE-SCREEN 2 CỘT) -----------------
col_left, col_right = st.columns([1, 1], gap="medium")

with col_left:
    st.markdown('<div class="section-title">📂 1. Chọn tệp & Thiết lập</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Kéo thả tệp PowerPoint (.pptx)", type=["pptx"], label_visibility="collapsed")
    
    if not uploaded_file and SAMPLE_PPTX.exists():
        if st.button("💡 Dùng thử nhanh file mẫu có sẵn"):
            with open(SAMPLE_PPTX, "rb") as sf:
                uploaded_file = io.BytesIO(sf.read())
                uploaded_file.name = "sample_presentation.pptx"
                st.session_state["use_sample"] = uploaded_file

    if "use_sample" in st.session_state and not uploaded_file:
        uploaded_file = st.session_state["use_sample"]

    if uploaded_file:
        st.success(f"📄 **{uploaded_file.name}** ({len(uploaded_file.getvalue())/1024:.1f} KB)")

    target_lang = st.selectbox(
        "🌐 Ngôn ngữ muốn dịch sang:",
        options=list(LANGS.keys()),
        format_func=lambda k: f"{LANGS[k][0]} ({k})"
    )

    with st.expander("📖 Bảng thuật ngữ tùy biến (Glossary)"):
        gl_text = st.text_area("Mỗi dòng: `thuật ngữ = bản dịch`", placeholder="API = Giao diện lập trình", height=60, label_visibility="collapsed")

    btn_disabled = not (uploaded_file and api_key)
    btn_translate = st.button("🚀 BẮT ĐẦU DỊCH POWERPOINT", type="primary", disabled=btn_disabled, use_container_width=True)

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
                tc1, tc2 = st.columns([2.3, 1], gap="small")
                with tc1:
                    lang_label = LANGS.get(t["target_lang"], (t["target_lang"],))[0]
                    st.markdown(
                        f"<div class='task-item-card'>"
                        f"📄 <b>{t['translated_file']}</b> "
                        f"<span class='task-size-tag'>{t['file_size_kb']} KB</span><br/>"
                        f"<span style='font-size: 0.75rem; opacity: 0.75;'>"
                        f"🌐 {lang_label} • ⏰ {t['created_at']}</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with tc2:
                    if t["file_exists"]:
                        with open(t["file_path"], "rb") as tf:
                            st.download_button(
                                label="📥 Tải file",
                                data=tf.read(),
                                file_name=t["translated_file"],
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key=f"dl_task_btn_{t['task_id']}_{idx}",
                                use_container_width=True
                            )
                    else:
                        st.caption("File đã xóa")
                st.markdown("<hr style='margin: 4px 0; opacity: 0.12;'/>", unsafe_allow_html=True)

    if btn_translate:
        st.session_state["live_logs"] = []
        st.session_state["translated_data"] = None
        render_logs(f"🚀 Bắt đầu tiếp nhận tệp: {uploaded_file.name}")
        
        def ui_progress(p, msg):
            progress_bar.progress(p)
            status_label.markdown(f"⏳ **Tiến độ {int(p * 100)}%:** {msg}")

        def ui_log(msg):
            render_logs(msg)

        try:
            with tempfile.TemporaryDirectory() as d:
                src_path = os.path.join(d, "input.pptx")
                dst_path = os.path.join(d, "translated.pptx")
                
                with open(src_path, "wb") as f_out:
                    f_out.write(uploaded_file.getvalue())
                
                ui_log(f"🤖 Khởi tạo bộ dịch Gemini AI: {model}")
                translator = GeminiTranslator(api_key=api_key, model=model, log_func=ui_log)
                glossary = load_glossary(gl_text) if gl_text.strip() else None
                
                report = translate_pptx(
                    src=src_path,
                    dst=dst_path,
                    target=target_lang,
                    translator=translator,
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
                
                with open(dst_path, "rb") as f_res:
                    st.session_state["translated_data"] = f_res.read()
                
                base_name = Path(uploaded_file.name).stem
                st.session_state["translated_filename"] = f"{base_name}.{target_lang}.pptx"
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
