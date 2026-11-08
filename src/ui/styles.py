"""CSS giao diện cao cấp: Sử dụng biến theme chuẩn của Streamlit, đảm bảo hiển thị 100% rõ nét ở cả Chế độ Sáng & Tối."""

SINGLE_PAGE_CSS = """
<style>
    /* =========================================================
       1. VIEWPORT CỐ ĐỊNH VỪA KHÍT 1 MÀN HÌNH (NO SCROLL NGOÀI)
       ========================================================= */
    html, body, [data-testid="stAppViewContainer"] {
        overflow-x: hidden !important;
    }

    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 1.8rem !important;
        padding-right: 1.8rem !important;
        max-width: 98% !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 1.8rem !important;
    }
    footer {
        visibility: hidden !important;
        display: none !important;
    }

    /* =========================================================
       2. TIÊU ĐỀ ỨNG DỤNG (ĐỒNG BỘ NATIVE STREAMLIT THEME)
       ========================================================= */
    .main-title {
        font-size: 1.35rem !important;
        font-weight: 800 !important;
        margin-bottom: 0.35rem !important;
        display: flex;
        align-items: center;
        gap: 0.6rem;
        letter-spacing: -0.02em;
    }
    .title-text {
        color: var(--text-color) !important;
        font-weight: 800 !important;
    }

    .main-sub {
        font-size: 0.84rem !important;
        font-weight: 500 !important;
        color: var(--text-color) !important;
        opacity: 0.75 !important;
        margin-left: 0.6rem !important;
    }

    .badge-tag {
        font-size: 0.72rem;
        font-weight: 700;
        padding: 0.15rem 0.55rem;
        border-radius: 9999px;
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
        color: #ffffff !important;
        border: 1px solid #38bdf8;
    }

    /* TIÊU ĐỀ TỪNG CỘT MỤC (LUÔN RÕ NÉT TRÊN MỌI THEME) */
    .section-title {
        font-size: 1.02rem !important;
        font-weight: 800 !important;
        color: var(--text-color) !important;
        margin-bottom: 0.3rem !important;
        display: flex;
        align-items: center;
        gap: 0.45rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    /* =========================================================
       3. CỐ ĐỊNH CHIỀU CAO LIVE TERMINAL LOG (220PX Ở CỘT TRÁI)
       ========================================================= */
    .terminal-container {
        border-radius: 12px !important;
        overflow: hidden !important;
        border: 1px solid #25334c !important;
        box-shadow: 0 4px 18px rgba(0, 0, 0, 0.25) !important;
        margin-top: 0.4rem !important;
    }

    .terminal-header {
        background: #0d131f !important;
        padding: 6px 12px !important;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #1e293b !important;
    }
    .terminal-header-dots {
        display: flex;
        gap: 6px;
    }
    .t-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
        display: inline-block;
    }
    .t-dot-red { background: #ff5f56; }
    .t-dot-yellow { background: #ffbd2e; }
    .t-dot-green { background: #27c93f; }
    .terminal-header-title {
        font-size: 0.74rem !important;
        font-weight: 700 !important;
        color: #e2e8f0 !important;
        letter-spacing: 0.06em !important;
        text-transform: uppercase !important;
        font-family: 'JetBrains Mono', 'Fira Code', monospace !important;
    }
    .terminal-header-badge {
        font-size: 0.68rem;
        font-weight: 600;
        color: #38bdf8;
        background: rgba(56, 189, 248, 0.15);
        padding: 1px 6px;
        border-radius: 4px;
    }

    .terminal-box {
        height: 220px !important;
        max-height: 220px !important;
        min-height: 220px !important;
        overflow-y: auto !important;
        background-color: #090d16 !important;
        padding: 10px 14px !important;
        font-family: 'JetBrains Mono', 'Fira Code', 'Consolas', monospace !important;
        font-size: 0.81rem !important;
        line-height: 1.6 !important;
        color: #e2e8f0 !important;
    }
    
    .terminal-box::-webkit-scrollbar {
        width: 6px !important;
    }
    .terminal-box::-webkit-scrollbar-track {
        background: #090d16 !important;
    }
    .terminal-box::-webkit-scrollbar-thumb {
        background: #38bdf8 !important;
        border-radius: 6px !important;
    }
    .terminal-box::-webkit-scrollbar-thumb:hover {
        background: #0ea5e9 !important;
    }

    /* =========================================================
       4. NÚT DỊCH CHÍNH (GRADIENT CAM-ĐỎ)
       ========================================================= */
    button[kind="primary"] {
        background: linear-gradient(135deg, #f97316 0%, #ef4444 100%) !important;
        color: #ffffff !important;
        font-size: 1rem !important;
        font-weight: 800 !important;
        padding: 0.6rem 1.2rem !important;
        border-radius: 12px !important;
        border: 1px solid #f87171 !important;
        box-shadow: 0 4px 18px rgba(239, 68, 68, 0.4) !important;
        transition: all 0.25s ease-in-out !important;
        width: 100% !important;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #fb923c 0%, #f43f5e 100%) !important;
        box-shadow: 0 6px 24px rgba(239, 68, 68, 0.6) !important;
        transform: translateY(-2px) !important;
    }

    /* =========================================================
       5. DANH SÁCH FILE ĐÃ DỊCH (SQLITE DB TASK CARDS)
       ========================================================= */
    .empty-tasks-card {
        background-color: var(--secondary-background-color) !important;
        border: 1px dashed rgba(128, 128, 128, 0.3) !important;
        border-radius: 12px !important;
        padding: 2.2rem 1.5rem !important;
        text-align: center !important;
        color: var(--text-color) !important;
        margin-top: 0.5rem !important;
    }

    .task-item-card {
        font-size: 0.85rem !important;
        color: var(--text-color) !important;
        line-height: 1.4 !important;
        padding: 2px 0 !important;
    }
    .task-size-tag {
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        padding: 1px 6px !important;
        border-radius: 4px !important;
        background-color: rgba(56, 189, 248, 0.15) !important;
        color: #0284c7 !important;
        margin-left: 4px !important;
    }

    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        color: #ffffff !important;
        font-size: 0.82rem !important;
        font-weight: 700 !important;
        padding: 0.35rem 0.85rem !important;
        border-radius: 8px !important;
        border: 1px solid #34d399 !important;
        box-shadow: 0 2px 8px rgba(16, 185, 129, 0.3) !important;
        transition: all 0.2s ease-in-out !important;
        width: 100% !important;
        min-height: 33px !important;
        height: 33px !important;
        margin: 0 !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #10b981 0%, #34d399 100%) !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.5) !important;
        transform: translateY(-1px) !important;
    }

    /* =========================================================
       6. BANNER MODEL ĐANG HOẠT ĐỘNG (SIDEBAR TOP)
       ========================================================= */
    .active-model-card {
        background: linear-gradient(135deg, rgba(2, 132, 199, 0.15) 0%, rgba(14, 165, 233, 0.06) 100%) !important;
        border: 1px solid rgba(56, 189, 248, 0.5) !important;
        border-radius: 12px !important;
        padding: 10px 14px !important;
        margin-bottom: 12px !important;
        box-shadow: 0 4px 14px rgba(2, 132, 199, 0.15) !important;
    }
    .active-model-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 4px;
    }
    .active-model-badge {
        font-size: 0.68rem !important;
        font-weight: 700 !important;
        color: #10b981 !important;
        background: rgba(16, 185, 129, 0.18) !important;
        padding: 2px 8px !important;
        border-radius: 9999px !important;
        border: 1px solid rgba(16, 185, 129, 0.4) !important;
        letter-spacing: 0.04em !important;
    }
    .active-model-title {
        font-size: 0.74rem !important;
        font-weight: 700 !important;
        color: var(--text-color) !important;
        opacity: 0.8 !important;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .active-model-name {
        font-size: 0.98rem !important;
        font-weight: 800 !important;
        color: #0284c7 !important;
        font-family: 'JetBrains Mono', 'Fira Code', monospace !important;
        letter-spacing: -0.01em !important;
    }
    .active-model-sub {
        font-size: 0.72rem !important;
        color: var(--text-color) !important;
        opacity: 0.7 !important;
        margin-top: 3px !important;
        line-height: 1.3 !important;
    }
    .quota-active-tag {
        font-size: 0.65rem !important;
        font-weight: 700 !important;
        background: #0284c7 !important;
        color: #ffffff !important;
        padding: 1px 6px !important;
        border-radius: 4px !important;
        margin-left: 6px !important;
        vertical-align: middle !important;
    }

    /* =========================================================
       7. KHUNG QUOTA CARD THÔNG MINH (SIDEBAR)
       ========================================================= */
    .quota-card {
        background-color: var(--secondary-background-color) !important;
        border: 1px solid rgba(128, 128, 128, 0.25) !important;
        border-radius: 12px !important;
        padding: 10px 12px !important;
        margin-bottom: 8px !important;
        font-size: 0.82rem !important;
        color: var(--text-color) !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08) !important;
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .quota-card:hover {
        transform: translateY(-1px);
    }
    .quota-card-ok {
        border-left: 4px solid #10b981 !important;
    }
    .quota-card-warn {
        border-left: 4px solid #f59e0b !important;
    }
    .quota-card-err {
        border-left: 4px solid #ef4444 !important;
    }

    .quota-header-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 4px;
    }
    .quota-model-name {
        font-weight: 700;
        font-size: 0.88rem;
        color: var(--text-color) !important;
    }
    .quota-ping-pill {
        font-size: 0.72rem;
        font-weight: 600;
        padding: 1px 7px;
        border-radius: 9999px;
        background: rgba(56, 189, 248, 0.15);
        color: #0284c7;
        border: 1px solid rgba(56, 189, 248, 0.3);
    }
    .quota-detail-line {
        margin: 2px 0;
        font-size: 0.78rem;
        color: var(--text-color) !important;
        opacity: 0.85;
    }
    .quota-detail-line b {
        color: var(--text-color) !important;
        opacity: 1;
    }

    .quota-usage-track {
        background-color: rgba(128, 128, 128, 0.2) !important;
        height: 6px;
        border-radius: 9999px;
        overflow: hidden;
        margin: 5px 0 6px 0;
    }
    .quota-usage-fill {
        height: 100%;
        border-radius: 9999px;
        transition: width 0.4s ease;
    }
    .fill-ok {
        background: linear-gradient(90deg, #10b981 0%, #34d399 100%) !important;
    }
    .fill-warn {
        background: linear-gradient(90deg, #f59e0b 0%, #ef4444 100%) !important;
    }

    .quota-reset-badge {
        display: block;
        margin-top: 4px;
        padding: 3px 7px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
        background: rgba(245, 158, 11, 0.12);
        color: #d97706;
        border: 1px solid rgba(245, 158, 11, 0.25);
    }
    .quota-reset-badge.ready {
        background: rgba(16, 185, 129, 0.12);
        color: #059669;
        border: 1px solid rgba(16, 185, 129, 0.25);
    }

    /* Dropdowns & File Uploader */
    section[data-testid="stFileUploader"] {
        padding: 0.1rem 0 !important;
    }
    section[data-testid="stFileUploader"] > div {
        padding: 0.6rem !important;
        border-radius: 12px !important;
    }

    div[data-baseweb="select"] > div {
        border-radius: 10px !important;
    }
</style>
"""
