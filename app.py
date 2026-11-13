import os
import copy
import re
import unicodedata
from datetime import datetime
from io import BytesIO

import openpyxl
from openpyxl.styles import PatternFill
import pandas as pd
import streamlit as st

from ai_assistant import (
    code_of, item_key,
    COLUMNS, KnowledgeStore, audit_quote, build_quote_rows, clean,
    detect_boq_columns, has_llm, sync_codes, find_header_row, find_layout,
    layout_columns, load_price_items, promote_header, suggest_matches, apply_approved,
    apply_cell_edits, remove_and_add_rows,
)
from io import BytesIO

# --- ĐƯỜNG DẪN (luôn tính theo thư mục chứa app.py) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "MauBaoGia.xlsx")
PRICE_LIST_PATH = os.path.join(BASE_DIR, "BangGia.xlsx")
KNOWLEDGE_PATH = os.path.join(BASE_DIR, "learned_knowledge.json")
PLACEHOLDER_SHEET = "XXXXXXXXX"
QUOTE_SHEET = "BÁO GIÁ"

# --- CẤU HÌNH TRANG & GIAO DIỆN ---
LOGO_PATH = os.path.join(BASE_DIR, "tsc_logo.png")

st.set_page_config(
    page_title="TSC - Hệ Thống Báo Giá",
    page_icon="🦅",
    layout="wide",
    initial_sidebar_state="expanded",
)

GOLD = "#C9A227"
GOLD_DEEP = "#9C7A12"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;600;700&display=swap');
    html, body, .stApp, [class*="css"] {{ font-family: 'Be Vietnam Pro', 'Segoe UI', Tahoma, sans-serif; }}
    .stApp {{ background: linear-gradient(180deg, #14161C 0%, #0E1014 100%); color: #ECE7D8; }}
    [data-testid="stSidebar"] {{ background: #181A21; border-right: 1px solid {GOLD_DEEP}; }}
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {{ color: {GOLD} !important; }}
    h1, h2, h3 {{ color: {GOLD} !important; font-weight: 700 !important; }}
    hr {{ border: none !important; height: 2px; background: linear-gradient(90deg, {GOLD}, transparent) !important; }}
    .stButton > button {{ background: linear-gradient(135deg, {GOLD}, {GOLD_DEEP}); color: #111 !important;
        border: none; border-radius: 8px; font-weight: 700; letter-spacing: .3px;
        transition: all .2s ease; box-shadow: 0 2px 8px rgba(201,162,39,.25); }}
    .stButton > button:hover {{ filter: brightness(1.08); box-shadow: 0 4px 14px rgba(201,162,39,.45); transform: translateY(-1px); }}
    .stAlert {{ background-color: #1E2028; border: 1px solid #3A3426; color: #ECE7D8; border-radius: 10px; }}
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {{ border-radius: 10px; overflow: hidden; border: 1px solid #3A3426; }}
    .streamlit-expanderHeader, [data-testid="stExpander"] summary {{ background-color: #1A1C23; color: {GOLD}; border-radius: 8px; }}
    [data-testid="stExpander"] {{ border: 1px solid #2E2A1E !important; border-radius: 10px; }}
    .brand-card {{ background: #FFFFFF; border-radius: 16px; padding: 18px 28px; display: flex; align-items: center;
        justify-content: center; box-shadow: 0 6px 24px rgba(0,0,0,.35); border-bottom: 4px solid {GOLD}; margin-bottom: 14px; }}
    .brand-card img {{ max-width: 520px; width: 100%; height: auto; }}
    .brand-sub {{ text-align: center; color: {GOLD}; font-size: 20px; font-weight: 600; letter-spacing: .5px; margin: 4px 0 6px; }}
    .gold-bar {{ height: 4px; border-radius: 4px; background: linear-gradient(90deg, {GOLD_DEEP}, {GOLD}, {GOLD_DEEP}); margin-bottom: 18px; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def logo_base64():
    import base64
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()


logo_b64 = logo_base64()
if logo_b64:
    st.markdown(f'<div class="brand-card"><img src="data:image/png;base64,{logo_b64}" alt="TSC"></div>',
                unsafe_allow_html=True)
else:
    st.markdown('<h1 style="text-align:center;">TSC TECHNICAL SOLUTION COMPANY LIMITED</h1>', unsafe_allow_html=True)
    st.warning("Chưa thấy file tsc_logo.png cạnh app.py, đang hiện chữ thay logo.")
st.markdown('<div class="brand-sub">HỆ THỐNG XỬ LÝ BÁO GIÁ THÔNG MINH</div>', unsafe_allow_html=True)
st.markdown('<div class="gold-bar"></div>', unsafe_allow_html=True)

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"❌ Không tìm thấy file mẫu `{TEMPLATE_PATH}`.")
    st.stop()

# --- TRẠNG THÁI PHIÊN ---
for _key, _default in {"df_editable": None, "learned_keys": set(), "last_stats": None, "editor_ver": 0}.items():
    if _key not in st.session_state:
        st.session_state[_key] = _default

store = KnowledgeStore(KNOWLEDGE_PATH)


@st.cache_data(show_spinner=False)
def cached_price_items(path, mtime, sheets):
    return load_price_items(path, list(sheets))


@st.cache_data(show_spinner=False)
def cached_price_sheet_names(path, mtime):
    return pd.ExcelFile(path).sheet_names


@st.cache_data(show_spinner=False)
def cached_boq_sheet_names(data: bytes):
    return pd.ExcelFile(BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False)
def cached_boq_raw(data: bytes, name: str, sheet):
    if name.lower().endswith(".csv"):
        return pd.read_csv(BytesIO(data), header=None)
    return pd.read_excel(BytesIO(data), sheet_name=sheet, header=None)


def on_editor_change(editor_key, price_items, store):
    """Chạy ngay khi người dùng sửa bảng: xóa dòng, đồng bộ mã, duyệt. Không cần bấm nút."""
    state = st.session_state.get(editor_key)
    if not isinstance(state, dict) or st.session_state.df_editable is None:
        return
    df = recalc(st.session_state.df_editable)
    edited_rows = state.get("edited_rows") or {}
    df = apply_cell_edits(df, edited_rows)
    synced = sync_codes(df, edited_rows, price_items, store, st.session_state.learned_keys)
    applied = apply_approved(df, price_items, store, st.session_state.learned_keys)
    df, structural = remove_and_add_rows(df, state)
    st.session_state.df_editable = recalc(df)
    if structural or synced or applied:
        st.session_state.editor_ver += 1


def ascii_key(v):
    """Chuỗi không dấu, chữ thường, để so khớp nhãn dù gõ sai dấu (ví dụ 'Thuể' / 'Thanh tóan')."""
    s = unicodedata.normalize("NFD", clean(v)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", s).strip()


def label_item(it):
    return f"[{it['ma']}] {it['ten']} - {it['don_gia']:,.0f} đ"


def recalc(df):
    df = df.copy()
    defaults = {"Xóa": False, "Duyệt": False, "SL": 0, "Đơn giá (VND)": 0,
                "Thành tiền (VND)": 0, "Độ khớp": 0.0}
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = defaults.get(c, "")
    df["Xóa"] = df["Xóa"].fillna(False).astype(bool)
    df["Duyệt"] = df["Duyệt"].fillna(False).astype(bool)
    df["Độ khớp"] = pd.to_numeric(df["Độ khớp"], errors="coerce").fillna(0.0)
    df["SL"] = pd.to_numeric(df["SL"], errors="coerce").fillna(0)
    df["Đơn giá (VND)"] = pd.to_numeric(df["Đơn giá (VND)"], errors="coerce").fillna(0)
    df["Thành tiền (VND)"] = df["SL"] * df["Đơn giá (VND)"]
    for c in ["Mã hàng", "Diễn giải", "Xuất xứ", "ĐVT", "Ghi chú", "BOQ gốc", "Gợi ý",
              "Mã đề xuất", "Nguồn"]:
        df[c] = df[c].fillna("")
    return df


def apply_item(df, idx, it):
    df.at[idx, "Mã hàng"] = item_key(it)
    df.at[idx, "Diễn giải"] = it["ten"]
    df.at[idx, "Xuất xứ"] = it["xuat_xu"]
    df.at[idx, "ĐVT"] = it["dvt"] or df.at[idx, "ĐVT"]
    df.at[idx, "Đơn giá (VND)"] = it["don_gia"]
    df.at[idx, "Nguồn"] = "Chọn tay"
    df.at[idx, "Duyệt"] = True


# --- SIDEBAR: TẢI DỮ LIỆU ---
st.sidebar.markdown(f"<h2 style='text-align: center; color: {GOLD};'>📥 TẢI DỮ LIỆU</h2>", unsafe_allow_html=True)
uploaded_boq = st.sidebar.file_uploader(
    "Tải lên File BOQ khối lượng gốc (Excel/CSV)", type=["xlsx", "xls", "csv"])

boq_sheet_selected = None
if uploaded_boq is not None and uploaded_boq.name.lower().endswith(("xlsx", "xls")):
    try:
        uploaded_boq.seek(0)
        xl_boq_names = cached_boq_sheet_names(uploaded_boq.getvalue())
        st.sidebar.markdown("---")
        boq_sheet_selected = st.sidebar.selectbox("📂 Chọn Sheet chứa khối lượng (BOQ):", xl_boq_names)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file BOQ: {e}")

price_items = []
if os.path.exists(PRICE_LIST_PATH):
    try:
        all_sheets = cached_price_sheet_names(PRICE_LIST_PATH, os.path.getmtime(PRICE_LIST_PATH))
        sheet_options = [s for s in all_sheets if s.upper() != PLACEHOLDER_SHEET]
        st.sidebar.markdown("---")
        price_sheets = st.sidebar.multiselect(
            "📋 Sheet Bảng Giá áp dụng:", sheet_options, default=sheet_options)
        if price_sheets:
            price_items = cached_price_items(
                PRICE_LIST_PATH, os.path.getmtime(PRICE_LIST_PATH), tuple(price_sheets))
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc Bảng Giá: {e}")

st.sidebar.markdown("---")
st.sidebar.markdown(f"<h3 style='color: {GOLD};'>🧠 TRỢ LÝ AI</h3>", unsafe_allow_html=True)
st.sidebar.caption(f"Bảng giá: {len(price_items)} mục · Đã học: {len(store.data['mappings'])} ghép mã")
use_llm = st.sidebar.checkbox(
    "Dùng Claude cho dòng khó", value=False, disabled=not has_llm(),
    help="Cần biến môi trường ANTHROPIC_API_KEY. Chỉ gọi cho các dòng hệ thống không chắc.")
if not has_llm():
    st.sidebar.caption("Chưa bật Claude (thiếu ANTHROPIC_API_KEY). Vẫn dùng gợi ý nội bộ.")

# --- ĐỌC FILE BOQ & NHẬN DIỆN CỘT ---
if uploaded_boq is None:
    st.info("👈 Vui lòng tải lên file Khối Lượng (BOQ) ở menu bên trái để bắt đầu lập báo giá.")
else:
    df_raw = None
    raw = None
    try:
        raw = cached_boq_raw(uploaded_boq.getvalue(), uploaded_boq.name, boq_sheet_selected).copy()
    except Exception as e:
        st.error(f"Lỗi đọc file BOQ: {e}")

    if raw is not None:
        raw = raw.dropna(how="all").reset_index(drop=True)
        if raw.empty:
            st.error(f"❌ Sheet '{boq_sheet_selected}' không có dữ liệu.")
        else:
            file_tag = f"{uploaded_boq.name}_{boq_sheet_selected}"
            layout, layout_row = find_layout(raw, store.data["layouts"])
            default_header = layout_row if layout else find_header_row(raw)
            labels = {"name": "Cột TÊN HÀNG", "qty": "Cột SỐ LƯỢNG (SL khách yêu cầu)",
                      "unit": "Cột ĐƠN VỊ", "spec": "Cột THÔNG SỐ / MÔ TẢ"}

            with st.expander("🔎 Xác nhận cách đọc file (lần đầu chọn, lần sau hệ thống tự nhớ)",
                             expanded=layout is None):
                if layout:
                    st.success("✅ Đã nhận ra cấu trúc file này. Cách chọn cột đã được nhớ từ lần trước.")
                else:
                    st.info("Cấu trúc file này chưa được nhớ. Hãy chọn đúng cột bên dưới. "
                            "Khi bấm xử lý, hệ thống sẽ ghi nhớ để lần sau tự chọn.")
                header_ui = st.number_input(
                    "Dòng tiêu đề (số dòng trong Excel)", min_value=1,
                    max_value=max(1, min(30, len(raw))), value=int(default_header) + 1,
                    key=f"hdr_{file_tag}")
                header_row = int(header_ui) - 1

                df_raw, _ = promote_header(raw, header_row)
                df_raw = df_raw.dropna(how="all")
                all_cols = [str(c) for c in df_raw.columns]
                heur = detect_boq_columns(df_raw)
                from_layout = layout_columns(df_raw, layout) if (layout and header_row == layout_row) else {}
                guess = {k: (from_layout.get(k) or heur.get(k)) for k in ["name", "qty", "unit", "spec", "price"]}

                opts = ["(không có)"] + all_cols
                c1, c2 = st.columns(2)
                cols = {}
                for i, key in enumerate(["name", "qty", "unit", "spec"]):
                    default = guess.get(key)
                    idx = opts.index(default) if default in opts else 0
                    chosen = (c1 if i % 2 == 0 else c2).selectbox(
                        labels[key], opts, index=idx, key=f"col_{key}_{file_tag}_{header_row}")
                    cols[key] = None if chosen == "(không có)" else chosen
                cols["price"] = guess.get("price")

                st.caption("Xem trước 8 dòng đầu của sheet:")
                st.dataframe(raw.head(8), use_container_width=True)

            if not cols["name"] or not cols["qty"]:
                st.error("❌ Chưa chọn đủ cột Tên hàng và Số lượng. Mở mục 'Xác nhận cách đọc file' để chọn.")
            elif st.sidebar.button("🚀 BẮT ĐẦU XỬ LÝ BOQ", type="primary", use_container_width=True):
                store.save_layout(raw, header_row, list(df_raw.columns), cols)
                if not price_items:
                    st.warning("Chưa có bảng giá nào được chọn ở sidebar. Các dòng sẽ không được ghép mã.")
                with st.spinner("AI đang phân tích và gợi ý mã hàng..."):
                    rows, stats = build_quote_rows(df_raw, cols, price_items, store, use_llm=use_llm)
                    st.session_state.df_editable = pd.DataFrame(rows, columns=COLUMNS)
                    st.session_state.learned_keys = set()
                    st.session_state.last_stats = stats
                    st.session_state.editor_ver += 1
                st.toast("💾 Đã ghi nhớ cách đọc file này cho lần sau")

# --- GIAO DIỆN CHỈNH SỬA ---
if st.session_state.df_editable is not None:
    df = recalc(st.session_state.df_editable)
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3>🛠 CHI TIẾT KHỐI LƯỢNG & ĐƠN GIÁ</h3>", unsafe_allow_html=True)

    stats = st.session_state.last_stats
    if stats:
        st.success(
            f"AI đã xử lý {stats['total']} hạng mục: {stats['learned']} hạng mục đã học sẵn (tích Duyệt sẵn), "
            f"{stats['proposed']} hạng mục có mã đề xuất, {stats['none']} hạng mục chưa có mã phù hợp. "
            f"Hãy kiểm tra cột 'Mã đề xuất' và tích cột 'Duyệt' để điền mã vào file.")

    # --- Tra cứu nhanh (có gợi ý và tự học) ---
    if price_items and len(df) > 0:
        with st.expander("🔍 TÌM KIẾM VÀ CHÈN MÃ NHANH VÀO DÒNG BẤT KỲ"):
            c1, c2 = st.columns([1, 2])
            with c1:
                row_idx = st.selectbox(
                    "1. Chọn dòng cần điền:", options=list(df.index), key="qp_row",
                    format_func=lambda x: f"Dòng {x + 1}: {clean(df.at[x, 'Diễn giải'])[:60]}")
            query = clean(df.at[row_idx, "BOQ gốc"]) or clean(df.at[row_idx, "Diễn giải"])
            suggestions = suggest_matches(query, price_items, store, top_k=5)
            options = {}
            for s in suggestions:
                options[f"★ {label_item(s['item'])}"] = s["item"]
            for it in price_items:
                if it["ma"]:
                    options.setdefault(label_item(it), it)
            with c2:
                choice = st.selectbox("2. Chọn thiết bị trong Bảng giá (★ = AI gợi ý):",
                                      ["-- Trống --"] + list(options), key="qp_item")
            if st.button("✅ ÁP DỤNG MÃ VÀO DÒNG ĐÃ CHỌN", key="qp_apply"):
                if choice != "-- Trống --":
                    it = options[choice]
                    before = clean(df.at[row_idx, "Mã hàng"])
                    apply_item(df, row_idx, it)
                    if query:
                        store.learn_mapping(query, it)
                        store.log_correction(query, before, item_key(it))
                    st.session_state.df_editable = recalc(df)
                    st.session_state.editor_ver += 1
                    st.rerun()

    # Điền mã cho các dòng đã tích Duyệt nhưng chưa được điền (kể cả khi bấm Duyệt tất cả)
    if apply_approved(df, price_items, store, st.session_state.learned_keys):
        st.session_state.df_editable = recalc(df)
        st.session_state.editor_ver += 1
        st.rerun()

    # --- Duyệt nhanh ---
    col_act1, _ = st.columns([1, 3])
    with col_act1:
        if st.button("✅ DUYỆT TẤT CẢ ĐỀ XUẤT", use_container_width=True):
            df.loc[df["Mã đề xuất"] != "", "Duyệt"] = True
            st.session_state.df_editable = df
            st.session_state.editor_ver += 1
            st.rerun()
    st.caption("Tích ô 🗑️ Xóa là dòng được xóa ngay. Tích ✅ Duyệt là mã đề xuất được điền vào file.")

    # --- Bảng chỉnh sửa (mọi thay đổi xử lý trong on_editor_change, không chạy lại cả trang) ---
    code_options = sorted({item_key(p) for p in price_items if p["ma"]})
    editor_key = f"boq_editor_{st.session_state.editor_ver}"
    st.data_editor(
        df,
        key=editor_key,
        num_rows="dynamic",
        use_container_width=True,
        height=500,
        on_change=on_editor_change,
        args=(editor_key, price_items, store),
        column_config={
            "Xóa": st.column_config.CheckboxColumn("🗑️ Xóa", help="Tích là xóa dòng ngay", default=False),
            "Duyệt": st.column_config.CheckboxColumn("✅ Duyệt", help="Tích để điền mã đề xuất vào file", default=False),
            "BOQ gốc": st.column_config.TextColumn("BOQ gốc", disabled=True, width="large"),
            "Mã đề xuất": st.column_config.TextColumn("Mã đề xuất (AI)", disabled=True),
            "Độ khớp": st.column_config.NumberColumn("Độ khớp", format="%.2f", disabled=True),
            "Nguồn": st.column_config.TextColumn("Nguồn", disabled=True),
            "Mã hàng": st.column_config.SelectboxColumn(
                "Mã hàng", options=code_options, help="Chọn mã + kích thước từ bảng giá. Tên và giá sẽ tự điền."),
            "Gợi ý": st.column_config.TextColumn("Gợi ý khác", disabled=True),
            "SL": st.column_config.NumberColumn("SL", format="%.2f", min_value=0),
            "Đơn giá (VND)": st.column_config.NumberColumn("Đơn giá (VND)", format="%,.0f", min_value=0),
            "Thành tiền (VND)": st.column_config.NumberColumn("Thành tiền (VND)", format="%,.0f", disabled=True),
        },
    )
    df = recalc(st.session_state.df_editable)

    # --- Kiểm tra trước khi xuất ---
    audit_rows = [{"ma": code_of(r["Mã hàng"]), "sl": r["SL"], "don_gia": r["Đơn giá (VND)"]}
                  for _, r in df.iterrows()]
    issues = audit_quote(audit_rows, store)
    n_err = sum(1 for i in issues if i["level"] == "error")
    n_warn = sum(1 for i in issues if i["level"] == "warn")
    with st.expander(f"🩺 KIỂM TRA BÁO GIÁ ({n_err} lỗi · {n_warn} cảnh báo)", expanded=n_err > 0):
        if not issues:
            st.success("Không phát hiện vấn đề.")
        show = {"error": st.error, "warn": st.warning, "info": st.info}
        for w in issues[:40]:
            show[w["level"]](w["message"])
        if len(issues) > 40:
            st.caption(f"... và {len(issues) - 40} cảnh báo khác.")

    # --- Thông tin khách hàng & điều khoản ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<h3>📝 THÔNG TIN DỰ ÁN & ĐIỀU KHOẢN THƯƠNG MẠI</h3>", unsafe_allow_html=True)

    with st.expander("👤 Thông tin Khách hàng (Header)", expanded=True):
        hc1, hc2 = st.columns(2)
        with hc1:
            h_to = st.text_input("To (Tên công ty):", placeholder="VD: Công ty TNHH ABC...", key="h_to")
            h_add = st.text_input("Add (Địa chỉ):", key="h_add")
            h_att = st.text_input("Att (Người liên hệ):", key="h_att")
            h_mobile = st.text_input("Mobile (Số điện thoại):", key="h_mobile")
            h_email = st.text_input("Email:", key="h_email")
            h_ref = st.text_input("Ref:", key="h_ref")
        with hc2:
            h_no = st.text_input("No (Số báo giá):", value=f"BG-{datetime.now().strftime('%d%m%y')}-01", key="h_no")
            h_date = st.date_input("Date (Ngày báo giá):", value=datetime.now(), key="h_date")
            h_revise = st.text_input("Revise (Lần chỉnh sửa):", value="0", key="h_revise")
            h_validity = st.text_input("Validity (Hiệu lực báo giá):", value="30 ngày", key="h_validity")
            h_project = st.text_input("Project (Tên dự án):", key="h_project")

    with st.expander("🤝 Điều khoản thương mại (Footer)", expanded=False):
        f_note = st.text_area("Lưu ý:", value="Báo giá đã bao gồm vận chuyển và giấy thông hành", key="f_note")
        f_deliv1 = st.text_input("Giao hàng (Dòng 1):", value="Tại địa chỉ dự án bên mua.", key="f_deliv1")
        f_deliv2 = st.text_input("Giao hàng (Dòng 2):",
                                 value="Nhập hàng 04-06 Tuần + giấy lưu thông 04-06 tuần", key="f_deliv2")
        f_payment = st.text_area("Thanh toán:",
                                 value="Tạm ứng 30% khi đặt hàng\nThanh toán 70% còn lại trước khi giao hàng",
                                 key="f_payment")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Xuất Excel (tìm dòng theo nhãn, tự thêm/bớt dòng, viết lại công thức tổng) ---
    def build_excel(data: pd.DataFrame) -> bytes:
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb[QUOTE_SHEET] if QUOTE_SHEET in wb.sheetnames else wb.active

        def safe_write(row, col, val):
            cell = ws.cell(row=row, column=col)
            for rng in ws.merged_cells.ranges:
                if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
                    cell = ws.cell(row=rng.min_row, column=rng.min_col)
                    break
            cell.value = val

        def find_row(prefixes, first, last, cols=(1, 2, 3)):
            for r in range(first, min(last, ws.max_row) + 1):
                for c in cols:
                    k = ascii_key(ws.cell(row=r, column=c).value)
                    if any(k.startswith(p) for p in prefixes):
                        return r
            return None

        def shift_rows(at_row, amount):
            """Thêm (amount > 0) hoặc xóa (amount < 0) dòng tại at_row, giữ merge và chiều cao dòng."""
            max_row = ws.max_row
            heights = {r: ws.row_dimensions[r].height for r in range(at_row, max_row + 1)
                       if r in ws.row_dimensions and ws.row_dimensions[r].height}
            deleted_upto = at_row - amount if amount < 0 else at_row
            merges = []
            for rng in list(ws.merged_cells.ranges):
                if rng.min_row >= at_row:
                    merges.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
                    ws.unmerge_cells(start_row=rng.min_row, start_column=rng.min_col,
                                     end_row=rng.max_row, end_column=rng.max_col)
            if amount > 0:
                ws.insert_rows(at_row, amount)
            else:
                ws.delete_rows(at_row, -amount)
            for r in range(at_row, max_row + abs(amount) + 1):
                if r in ws.row_dimensions:
                    ws.row_dimensions[r].height = None
            for r, h in heights.items():
                if r >= deleted_upto:
                    ws.row_dimensions[r + amount].height = h
            for r1, c1, r2, c2 in merges:
                if r1 < deleted_upto:
                    continue
                ws.merge_cells(start_row=r1 + amount, start_column=c1,
                               end_row=r2 + amount, end_column=c2)

        # 1) Header (vị trí giống mẫu Q1011 và mẫu hiện tại)
        safe_write(6, 2, h_to)
        safe_write(7, 2, h_add)
        safe_write(8, 2, h_att)
        safe_write(9, 2, h_mobile)
        safe_write(10, 2, h_email)
        safe_write(11, 2, h_ref)
        safe_write(6, 7, h_no)
        safe_write(7, 7, h_date.strftime("%d/%m/%Y"))
        safe_write(8, 7, h_revise)
        safe_write(9, 7, h_validity)
        safe_write(10, 7, h_project)

        # 2) Tìm dòng tiêu đề "Mã hàng" và dòng "Cộng" làm mốc vùng hàng hóa
        hdr = find_row(["ma hang"], 1, 60, cols=(2,))
        if hdr is None:
            raise ValueError("Không tìm thấy dòng tiêu đề 'Mã hàng' trong mẫu báo giá.")
        start_row = hdr + 1
        total_row = find_row(["cong"], start_row, start_row + 500)
        if total_row is None:
            raise ValueError("Không tìm thấy dòng 'Cộng' trong mẫu báo giá.")
        available = total_row - start_row
        n = len(data)
        last_item_row = total_row - 1

        # 3) Thêm hoặc bớt dòng hàng hóa cho khớp số hàng
        if n > available:
            extra = n - available
            shift_rows(total_row, extra)
            for r in range(total_row, total_row + extra):
                for col in range(1, 10):
                    s_cell = ws.cell(row=last_item_row, column=col)
                    t_cell = ws.cell(row=r, column=col)
                    if s_cell.has_style:
                        t_cell.font = copy.copy(s_cell.font)
                        t_cell.border = copy.copy(s_cell.border)
                        t_cell.fill = copy.copy(s_cell.fill)
                        t_cell.number_format = s_cell.number_format
                        t_cell.alignment = copy.copy(s_cell.alignment)
        elif n < available:
            fewer = available - n
            shift_rows(start_row + n, -fewer)

        # 4) Ghi hàng hóa
        flag_fill = PatternFill("solid", fgColor="FFF2A8")
        for i, (_, r) in enumerate(data.iterrows()):
            rr = start_row + i
            code = code_of(r["Mã hàng"])
            price = float(r["Đơn giá (VND)"] or 0)
            name = clean(r["Diễn giải"]) if code else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
            notes = []
            if clean(r["Ghi chú"]):
                notes.append(clean(r["Ghi chú"]))
            if not code:
                notes.append("CHƯA GHÉP MÃ")
            if price <= 0:
                notes.append("CHƯA CÓ GIÁ")
            safe_write(rr, 1, i + 1)
            safe_write(rr, 2, code)
            safe_write(rr, 3, name)
            safe_write(rr, 4, clean(r["Xuất xứ"]))
            safe_write(rr, 5, clean(r["ĐVT"]) or "Cái")
            safe_write(rr, 6, float(r["SL"]))
            safe_write(rr, 7, price if price > 0 else None)
            safe_write(rr, 8, f"=F{rr}*G{rr}")
            safe_write(rr, 9, " | ".join(notes))
            if not code or price <= 0:
                for c in range(1, 10):
                    ws.cell(row=rr, column=c).fill = flag_fill

        # 5) Công thức Cộng, Thuế, Tổng cộng: viết lại theo vị trí mới
        cong = start_row + n
        safe_write(cong, 8, f"=SUM(H{start_row}:H{cong - 1})")
        thue = find_row(["thue"], cong + 1, cong + 3)
        tong = find_row(["tong cong"], cong + 1, cong + 6)
        if thue is not None:
            rate = ws.cell(row=thue, column=4).value
            if isinstance(rate, (int, float)):
                safe_write(thue, 8, f"=H{cong}*D{thue}")
            else:
                safe_write(thue, 4, 0.08)
                safe_write(thue, 8, f"=H{cong}*D{thue}")
        if tong is not None:
            safe_write(tong, 8, f"=H{cong}+H{thue}" if thue is not None else f"=H{cong}")

        # 6) Điều khoản dưới bảng
        footer_start = (tong or cong) + 1
        for r in range(footer_start, ws.max_row + 1):
            k2 = ascii_key(ws.cell(row=r, column=2).value)
            k3 = ascii_key(ws.cell(row=r, column=3).value)
            if k2.startswith("luu y"):
                safe_write(r, 3, f_note)
            elif k2.startswith("giao hang"):
                safe_write(r, 3, f_deliv1)
            elif k3.startswith("nhap hang") or "giay luu thong" in k3:
                safe_write(r, 3, f_deliv2)
            elif k2.startswith("thanh toan"):
                lines = [ln.strip() for ln in f_payment.split("\n") if ln.strip()]
                slots = [r]
                rr = r + 1
                while rr <= ws.max_row and len(slots) < max(1, len(lines)):
                    if ascii_key(ws.cell(row=rr, column=3).value).startswith("thanh toan bang") \
                            or clean(ws.cell(row=rr, column=2).value):
                        break
                    slots.append(rr)
                    rr += 1
                for idx_s, srow in enumerate(slots):
                    if idx_s < len(lines) - 1:
                        safe_write(srow, 3, lines[idx_s])
                    elif idx_s == len(slots) - 1:
                        safe_write(srow, 3, " / ".join(lines[idx_s:]) if lines else None)
                    else:
                        safe_write(srow, 3, None)
                break

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    output_filename = st.text_input(
        "Tên file Excel xuất ra:", value=f"BaoGia_TSC_{datetime.now().strftime('%Y%m%d')}.xlsx", key="out_name")

    if st.button("📥 XUẤT FILE BÁO GIÁ", use_container_width=True):
        n_nocode = int(sum(1 for r in audit_rows if not r["ma"]))
        n_noprice = int(sum(1 for r in audit_rows if not r["don_gia"]))
        if n_nocode or n_noprice:
            st.info(f"File vẫn được xuất. {n_nocode} dòng chưa ghép mã (ghi tên BOQ gốc), "
                    f"{n_noprice} dòng chưa có giá (để trống). Các dòng này được tô vàng và ghi chú để bạn soát lại.")
        try:
            with st.spinner("Đang ghi dữ liệu vào mẫu báo giá..."):
                excel_bytes = build_excel(df)
            store.record_quote_prices(audit_rows)
            st.success("🎉 GHI DỮ LIỆU THÀNH CÔNG! Bấm tải xuống bên dưới.")
            st.download_button(
                label="⬇️ NHẤN ĐỂ LƯU FILE EXCEL",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Đã xảy ra lỗi khi xuất file: {e}")
