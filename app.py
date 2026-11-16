"""
app.py - Giao diện mới: menu bên trái, màn hình Báo giá mới 3 bước, Lịch sử, Danh mục.
Xử lý nghiệp vụ nằm trong quote_core.py, ai_assistant.py, lich_su.py.
"""
import base64
import os
from datetime import datetime

import pandas as pd
import streamlit as st

import lich_su
from ai_assistant import (
    COLUMNS, KnowledgeStore, apply_approved, apply_cell_edits, audit_quote, build_quote_rows,
    clean, code_of, detect_boq_columns, find_header_row, find_layout, has_llm, item_key,
    layout_columns, load_price_items, promote_header, remove_and_add_rows, save_override, suggest_matches,
    sync_codes,
)
from quote_core import (
    LOGO_PATH, QUOTE_SHEET, TEMPLATE_PATH, build_excel, build_quote_pdf, quote_totals, recalc,
    render_preview_html, xlsx_to_pdf,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KNOWLEDGE_PATH = os.path.join(BASE_DIR, "learned_knowledge.json")
PLACEHOLDER_SHEET = "XXXXXXXXX"
PRICE_PATH = os.path.join(BASE_DIR, "BangGia.xlsx")
PRIMARY, ACCENT = "#1F4E79", "#2E75B6"

st.set_page_config(page_title="TSC - Phần mềm báo giá", page_icon="🦅", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown(f"""
<style>
    html, body, .stApp, [class*="css"] {{ font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 14px; }}
    .stApp {{ background: #EEF2F6; color: #1F2937; }}
    .block-container {{ padding-top: 0.8rem; max-width: 100%; }}
    [data-testid="stSidebar"] {{ background: {PRIMARY}; }}
    [data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
    [data-testid="stSidebar"] [role="radiogroup"] label {{ padding: 6px 8px; border-radius: 4px; }}
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {{ background: rgba(255,255,255,.12); }}
    h1, h2, h3 {{ color: {PRIMARY} !important; font-weight: 700 !important; }}
    .topbar {{ background: {PRIMARY}; color: #fff; padding: 8px 16px; border-radius: 4px; margin-bottom: 10px;
        display: flex; align-items: center; gap: 14px; }}
    .topbar .t {{ font-size: 16px; font-weight: 700; }}
    .topbar .s {{ font-size: 13px; opacity: .85; margin-left: auto; }}
    .statusbar {{ background: #fff; border: 1px solid #C9D3DE; border-top: 2px solid {ACCENT}; border-radius: 4px;
        padding: 7px 14px; display: flex; flex-wrap: wrap; gap: 26px; font-size: 13px; color: #374151; margin: 4px 0 10px; }}
    .statusbar b {{ color: {PRIMARY}; }}
    .stButton > button {{ background: #FFFFFF; color: {PRIMARY} !important; border: 1px solid #9FB3C8;
        border-radius: 4px; font-weight: 600; font-size: 13px; min-height: 34px; }}
    .stButton > button:hover {{ background: #E8F0F8; border-color: {ACCENT}; }}
    .stButton > button[kind="primary"] {{ background: {ACCENT}; color: #fff !important; border-color: {ACCENT}; }}
    [data-testid="stExpander"] {{ background: #FFFFFF; border: 1px solid #C9D3DE !important; border-radius: 4px; }}
    [data-testid="stDataEditor"] {{ border: 1px solid #9FB3C8; border-radius: 4px; background: #FFFFFF; }}
    .stTabs [data-baseweb="tab"] {{ font-weight: 600; color: {PRIMARY}; }}

    /* ---- Vùng nội dung chính: chữ tối trên nền sáng, bất kể giao diện nền của Streamlit ---- */
    .block-container, .block-container p, .block-container span, .block-container label,
    .block-container div, .block-container li, .block-container h4, .block-container h5 {{ color: #1F2937; }}
    .block-container [data-testid="stWidgetLabel"] p, .block-container [data-testid="stWidgetLabel"] label {{
        color: #1F2937 !important; font-weight: 600; }}
    .block-container [data-testid="stCaptionContainer"], .block-container .stCaption {{ color: #4B5563 !important; }}
    .block-container [data-baseweb="select"] > div, .block-container [data-baseweb="input"],
    .block-container [data-baseweb="textarea"], .block-container textarea, .block-container input,
    .block-container [data-testid="stFileUploader"] section, .block-container [data-testid="stNumberInput"] input,
    .block-container [data-testid="stDateInput"] input {{
        background: #FFFFFF !important; color: #1F2937 !important; border-color: #9FB3C8 !important; }}
    .block-container [data-baseweb="select"] span, .block-container [data-baseweb="tag"] span {{ color: #1F2937; }}
    .block-container [data-baseweb="tag"] {{ background: #E8F0F8 !important; }}
    .block-container [data-testid="stFileUploader"] small, .block-container [data-testid="stFileUploader"] span,
    .block-container [data-testid="stFileUploader"] button {{ color: #1F2937 !important; }}
    .block-container [data-testid="stExpander"] summary p {{ color: {PRIMARY} !important; font-weight: 600; }}
    .block-container [data-testid="stAlert"] p, .block-container .stAlert p {{ color: #1F2937 !important; }}
    .block-container .stButton > button[kind="primary"] {{ color: #FFFFFF !important; }}
    .block-container .stButton > button[kind="primary"] p {{ color: #FFFFFF !important; }}
    .block-container .stButton > button p {{ color: {PRIMARY}; }}
    .block-container .stTabs [data-baseweb="tab"] p {{ color: {PRIMARY} !important; }}

    /* Khung tải file: nền sáng, chữ tối */
    .block-container [data-testid="stFileUploaderDropzone"],
    .block-container [data-testid="stFileUploader"] section,
    .block-container [data-testid="stFileUploaderFile"],
    .block-container [data-testid="stFileUploaderFileData"] {{
        background: #FFFFFF !important; border: 1px dashed #9FB3C8 !important; color: #1F2937 !important; }}
    .block-container [data-testid="stFileUploaderFileName"],
    .block-container [data-testid="stFileUploaderFileData"] span,
    .block-container [data-testid="stFileUploaderFileData"] small,
    .block-container [data-testid="stFileUploaderDropzoneInstructions"] span,
    .block-container [data-testid="stFileUploaderDropzoneInstructions"] small {{ color: #1F2937 !important; }}
    .block-container [data-testid="stFileUploaderDeleteBtn"] button,
    .block-container [data-testid="stFileUploader"] button {{ background: #FFFFFF !important; color: {PRIMARY} !important;
        border: 1px solid #9FB3C8 !important; }}
    .block-container [data-testid="stFileUploader"] svg {{ fill: {PRIMARY} !important; color: {PRIMARY} !important; }}

    /* Danh sách thả xuống (popup nằm ngoài vùng nội dung) */
    [data-baseweb="popover"] [role="listbox"], [data-baseweb="menu"], [data-baseweb="popover"] ul {{
        background: #FFFFFF !important; }}
    [data-baseweb="popover"] li, [data-baseweb="menu"] li, [role="option"] {{ color: #1F2937 !important; background: #FFFFFF; }}
    [data-baseweb="popover"] li:hover, [role="option"]:hover {{ background: #E8F0F8 !important; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def logo_b64():
    if not os.path.exists(LOGO_PATH):
        return None
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()


@st.cache_data(show_spinner=False)
def cached_price_items(path, mtime, sheets):
    return load_price_items(path, list(sheets))


@st.cache_data(show_spinner=False)
def cached_price_sheet_names(path, mtime):
    return pd.ExcelFile(path).sheet_names


@st.cache_data(show_spinner=False)
def cached_boq_sheet_names(data: bytes):
    from io import BytesIO
    return pd.ExcelFile(BytesIO(data)).sheet_names


@st.cache_data(show_spinner=False)
def cached_boq_raw(data: bytes, name: str, sheet):
    from io import BytesIO
    if name.lower().endswith(".csv"):
        return pd.read_csv(BytesIO(data), header=None)
    return pd.read_excel(BytesIO(data), sheet_name=sheet, header=None)


def get_store():
    return KnowledgeStore(KNOWLEDGE_PATH)


def get_price_items():
    """Bảng giá đang áp dụng (theo các sheet người dùng chọn ở bước 1)."""
    sheets = st.session_state.get("price_sheets")
    if not os.path.exists(PRICE_PATH) or not sheets:
        return []
    return cached_price_items(PRICE_PATH, os.path.getmtime(PRICE_PATH), tuple(sheets))


def label_item(it):
    return f"[{it['ma']}] {it['ten']} - {it['don_gia']:,.0f} đ"


MOBILE_CSS = """
<style>
    [data-testid="column"] {{ width: 100% !important; flex: 1 1 100% !important; min-width: 100% !important; }}
    .block-container {{ padding: .5rem .6rem !important; }}
    .stButton > button {{ min-height: 48px !important; font-size: 16px !important; width: 100%; }}
    .topbar .t {{ font-size: 14px; }}
    .statusbar {{ font-size: 12px; gap: 10px; }}
</style>
"""

if "device" not in st.session_state:
    st.markdown("<h2 style='text-align:center;'>Bạn đang dùng thiết bị nào?</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;'>App sẽ hiển thị phù hợp với thiết bị của bạn.</p>",
                unsafe_allow_html=True)
    d1, d2 = st.columns(2)
    if d1.button("💻 Máy tính", use_container_width=True, key="dev_pc"):
        st.session_state["device"] = "pc"
        st.rerun()
    if d2.button("📱 Điện thoại", use_container_width=True, key="dev_phone"):
        st.session_state["device"] = "mobile"
        st.rerun()
    st.stop()

IS_MOBILE = st.session_state["device"] == "mobile"
if IS_MOBILE:
    st.markdown(MOBILE_CSS, unsafe_allow_html=True)

PAGES = ["📝 Báo giá mới", "🗂️ Lịch sử báo giá", "📦 Danh mục hàng hóa"]

# ----------------------------- Thanh bên & menu -----------------------------
if IS_MOBILE:
    page = st.radio("Chức năng", PAGES, horizontal=False, key="page_m", label_visibility="collapsed")
    if st.button("🔄 Đổi thiết bị", key="dev_change_m"):
        del st.session_state["device"]
        st.rerun()
else:
  with st.sidebar:
    lb = logo_b64()
    if lb:
        st.markdown(f'<div style="background:#fff;border-radius:4px;padding:6px;text-align:center;">'
                    f'<img src="data:image/png;base64,{lb}" style="max-width:100%;height:auto;"></div>',
                    unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    PAGES = ["📝 Báo giá mới", "🗂️ Lịch sử báo giá", "📦 Danh mục hàng hóa"]
    page = st.radio("Chức năng", PAGES, key="page", label_visibility="collapsed")
    st.markdown("---")
    st.caption("Dữ liệu học, Cài đặt sẽ có ở đợt tiếp theo.")

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"❌ Không tìm thấy file mẫu `{TEMPLATE_PATH}`.")
    st.stop()

st.markdown(f'<div class="topbar"><span class="t">PHẦN MỀM BÁO GIÁ TSC</span>'
            f'<span class="s">{page.split(" ", 1)[1]}</span></div>', unsafe_allow_html=True)


# =============================== BÁO GIÁ MỚI ===============================
def page_quote():
    if "editor_ver" not in st.session_state:
        st.session_state.editor_ver = 0
        st.session_state.df_quote = None
        st.session_state.learned_keys = set()
        st.session_state.last_stats = None

    STEPS = ["1 · Tải BOQ", "2 · Ghép mã", "3 · Xuất báo giá"]
    target = st.session_state.pop("goto_step", None)
    if target:
        st.session_state["step"] = target
    step = st.radio("Bước", STEPS, horizontal=True, key="step", label_visibility="collapsed")

    # ---------- Bước 1 ----------
    if step == STEPS[0]:
        up = st.file_uploader("Chọn file BOQ khối lượng (Excel hoặc CSV)", type=["xlsx", "xls", "csv"],
                              key="boq_file")
        if up is None:
            st.info("Tải file BOQ để bắt đầu. Xong bước này app sẽ tự chuyển sang bước 2.")
        else:
            data = up.getvalue()
            sheet = None
            if not up.name.lower().endswith(".csv"):
                sheet = st.selectbox("Sheet chứa khối lượng", cached_boq_sheet_names(data), key="boq_sheet")
            store = get_store()
            try:
                raw = cached_boq_raw(data, up.name, sheet).copy().dropna(how="all").reset_index(drop=True)
            except Exception as e:
                st.error(f"Không đọc được file: {e}")
                raw = None
            if raw is not None and not raw.empty:
                layout, layout_row = find_layout(raw, store.data["layouts"])
                default_hdr = layout_row if layout else find_header_row(raw)
                tag = f"{up.name}_{sheet}"
                if layout:
                    st.success("✅ Đã nhận ra cấu trúc file này, cách chọn cột lấy từ lần trước.")
                hdr = st.number_input("Dòng tiêu đề (số dòng trong Excel)", 1, max(1, min(30, len(raw))),
                                      int(default_hdr) + 1, key=f"hdr_{tag}")
                header_row = int(hdr) - 1
                df_raw, _ = promote_header(raw, header_row)
                df_raw = df_raw.dropna(how="all")
                all_cols = [str(c) for c in df_raw.columns]
                heur = detect_boq_columns(df_raw)
                lay = layout_columns(df_raw, layout) if (layout and header_row == layout_row) else {}
                guess = {k: (lay.get(k) or heur.get(k)) for k in ["name", "qty", "unit", "spec"]}
                opts = ["(không có)"] + all_cols
                c1, c2, c3, c4 = st.columns(4)
                labels = {"name": ("Cột tên hàng", c1), "qty": ("Cột số lượng (SL)", c2),
                          "unit": ("Cột đơn vị", c3), "spec": ("Cột thông số/mô tả", c4)}
                cols = {}
                for key, (lab, col) in labels.items():
                    d = guess.get(key)
                    idx = opts.index(d) if d in opts else 0
                    ch = col.selectbox(lab, opts, index=idx, key=f"col_{key}_{tag}_{header_row}")
                    cols[key] = None if ch == "(không có)" else ch
                cols["price"] = guess.get("price")
                st.caption("8 dòng đầu của sheet:")
                st.dataframe(raw.head(8), use_container_width=True)

                names = [s for s in cached_price_sheet_names(PRICE_PATH, os.path.getmtime(PRICE_PATH))
                         if s.upper() != PLACEHOLDER_SHEET] if os.path.exists(PRICE_PATH) else []
                st.session_state.price_sheets = st.multiselect(
                    "Sheet bảng giá áp dụng", names,
                    default=st.session_state.get("price_sheets") or names, key="price_pick")

                if not cols["name"] or not cols["qty"]:
                    st.error("Chưa chọn đủ cột tên hàng và số lượng.")
                elif st.button("Tiếp: ghép mã  ➜", type="primary", key="btn_build"):
                    store.save_layout(raw, header_row, list(df_raw.columns), cols)
                    with st.spinner("Đang gợi ý mã hàng..."):
                        rows, stats = build_quote_rows(df_raw, cols, get_price_items(), store,
                                                       use_llm=False)
                    st.session_state.df_quote = pd.DataFrame(rows, columns=COLUMNS)
                    st.session_state.learned_keys = set()
                    st.session_state.last_stats = stats
                    st.session_state.editor_ver += 1
                    st.session_state["goto_step"] = STEPS[1]
                    st.rerun()

    # ---------- Bước 2 ----------
    if step == STEPS[1]:
        if st.session_state.df_quote is None:
            st.info("Chưa có dữ liệu. Làm bước 1 trước.")
        else:
            page_match()
            st.markdown("---")
            n1, n2, _ = st.columns([1, 1, 3])
            if n1.button("⬅ Quay lại bước 1", key="back_1"):
                go_step("1 · Tải BOQ")
            if n2.button("Tiếp: xuất báo giá ➜", type="primary", key="next_3"):
                go_step("3 · Xuất báo giá")

    # ---------- Bước 3 ----------
    if step == STEPS[2]:
        if st.session_state.df_quote is None:
            st.info("Chưa có dữ liệu. Làm bước 1 và 2 trước.")
        else:
            if st.button("⬅ Quay lại bước 2", key="back_2"):
                go_step("2 · Ghép mã")
            page_export()


def on_editor_change(editor_key, price_items, store):
    state = st.session_state.get(editor_key)
    if not isinstance(state, dict) or st.session_state.df_quote is None:
        return
    df = recalc(st.session_state.df_quote)
    edited_rows = state.get("edited_rows") or {}
    df = apply_cell_edits(df, edited_rows)
    synced = sync_codes(df, edited_rows, price_items, store, st.session_state.learned_keys)
    applied = apply_approved(df, price_items, store, st.session_state.learned_keys)
    df, structural = remove_and_add_rows(df, state)
    st.session_state.df_quote = recalc(df)
    if structural or synced or applied:
        st.session_state.editor_ver += 1


def go_step(name):
    st.session_state["goto_step"] = name
    st.rerun()


def page_match():
    price_items = get_price_items()
    store = get_store()
    df = recalc(st.session_state.df_quote)
    st.session_state.df_quote = df

    if not price_items:
        st.warning("Chưa chọn sheet bảng giá ở bước 1.")

    n_rows = len(df)
    total = quote_totals(df)[2]
    no_code = int(sum(1 for _, r in df.iterrows() if not code_of(r["Mã hàng"])))
    no_price = int(sum(1 for _, r in df.iterrows() if float(r["Đơn giá (VND)"] or 0) <= 0))
    st.markdown(
        f'<div class="statusbar"><span>Số dòng: <b>{n_rows}</b></span>'
        f'<span>Tổng tiền (có VAT): <b>{total:,.0f} VND</b></span>'
        f'<span>Chưa ghép mã: <b>{no_code}</b></span><span>Chưa có giá: <b>{no_price}</b></span>'
        f'<span>Đã học: <b>{len(store.data["mappings"])}</b> ghép mã</span></div>', unsafe_allow_html=True)

    with st.expander("🔍 Tìm nhanh mã để gán vào một dòng"):
        if price_items and n_rows:
            c1, c2 = st.columns([1, 2])
            row_idx = c1.selectbox("Dòng", list(df.index), key="qp_row",
                                   format_func=lambda x: f"Dòng {x + 1}: {clean(df.at[x, 'Diễn giải'])[:60]}")
            query = clean(df.at[row_idx, "BOQ gốc"]) or clean(df.at[row_idx, "Diễn giải"])
            opts = {}
            for s in suggest_matches(query, price_items, store, top_k=5):
                opts[f"★ {label_item(s['item'])}"] = s["item"]
            for it in price_items:
                if it["ma"]:
                    opts.setdefault(label_item(it), it)
            choice = c2.selectbox("Mã (★ = gợi ý)", ["-- Chọn --"] + list(opts), key="qp_item")
            if st.button("Gán mã vào dòng", key="qp_apply") and choice != "-- Chọn --":
                it = opts[choice]
                from quote_core import apply_item
                apply_item(df, row_idx, it)
                if query:
                    store.learn_mapping(query, it)
                    store.log_correction(query, "", item_key(it))
                st.session_state.df_quote = recalc(df)
                st.session_state.editor_ver += 1
                st.rerun()

    c_a, c_b, _ = st.columns([1, 1, 3])
    if c_a.button("✅ Duyệt tất cả đề xuất", key="btn_approve_all"):
        df.loc[df["Mã đề xuất"] != "", "Duyệt"] = True
        st.session_state.df_quote = df
        st.session_state.editor_ver += 1
        st.rerun()
    if apply_approved(df, price_items, store, st.session_state.learned_keys):
        st.session_state.df_quote = recalc(df)
        st.session_state.editor_ver += 1
        st.rerun()

    key = f"boq_editor_{st.session_state.editor_ver}"
    st.caption("Tích ✅ Duyệt để điền mã đề xuất. Tích 🗑️ Xóa để xóa dòng ngay. "
               "Chọn mã ở cột Mã hàng thì tên và giá tự điền.")
    code_options = sorted({item_key(p) for p in price_items if p["ma"]})
    st.data_editor(
        df, key=key, num_rows="dynamic", use_container_width=True, height=460,
        on_change=on_editor_change, args=(key, price_items, store),
        column_config={
            "Xóa": st.column_config.CheckboxColumn("🗑️ Xóa", default=False),
            "Duyệt": st.column_config.CheckboxColumn("✅ Duyệt", default=False),
            "BOQ gốc": st.column_config.TextColumn("BOQ gốc", disabled=True, width="large"),
            "Mã đề xuất": st.column_config.TextColumn("Mã đề xuất (AI)", disabled=True),
            "Độ khớp": st.column_config.NumberColumn("Độ khớp", format="%.2f", disabled=True),
            "Nguồn": st.column_config.TextColumn("Nguồn", disabled=True),
            "Mã hàng": st.column_config.SelectboxColumn("Mã hàng", options=code_options),
            "Gợi ý": st.column_config.TextColumn("Gợi ý khác", disabled=True),
            "SL": st.column_config.NumberColumn("SL", format="%.2f", min_value=0),
            "Đơn giá (VND)": st.column_config.NumberColumn("Đơn giá (VND)", format="%,.0f", min_value=0),
            "Thành tiền (VND)": st.column_config.NumberColumn("Thành tiền (VND)", format="%,.0f", disabled=True),
        })


def page_export():
    df = recalc(st.session_state.df_quote)
    st.markdown("#### Thông tin khách hàng")
    a1, a2, a3 = st.columns(3)
    H = {
        "to": a1.text_input("Khách hàng (To)", key="h_to"),
        "add": a1.text_input("Địa chỉ (Add)", key="h_add"),
        "att": a1.text_input("Người liên hệ (Att)", key="h_att"),
        "mobile": a2.text_input("Điện thoại (Mobile)", key="h_mobile"),
        "email": a2.text_input("Email", key="h_email"),
        "ref": a2.text_input("Tham chiếu (Ref)", key="h_ref"),
        "no": a3.text_input("Số báo giá (No)", value=f"BG-{datetime.now().strftime('%d%m%y')}-01", key="h_no"),
        "date": a3.date_input("Ngày (Date)", value=datetime.now(), key="h_date"),
        "revise": a3.text_input("Lần sửa (Revise)", value="0", key="h_revise"),
        "validity": a3.text_input("Hiệu lực (Validity)", value="30 ngày", key="h_validity"),
        "project": a3.text_input("Dự án (Project)", key="h_project"),
    }
    with st.expander("Điều khoản thương mại"):
        F = {
            "note": st.text_area("Lưu ý", value="Báo giá đã bao gồm vận chuyển và giấy thông hành", key="f_note"),
            "deliv1": st.text_input("Giao hàng", value="Tại địa chỉ dự án bên mua.", key="f_deliv1"),
            "deliv2": st.text_input("Giao hàng (dòng 2)", value="", key="f_deliv2"),
            "payment": st.text_area("Thanh toán", value="Tạm ứng 30% khi đặt hàng\nThanh toán 70% còn lại trước khi giao hàng",
                                    key="f_payment"),
            "bank": st.text_area("Thông tin tài khoản (nếu cần in thêm)", value="", key="f_bank"),
        }

    st.markdown("#### Xem trước")
    hv = dict(H, date=H["date"].strftime("%d/%m/%Y"))
    with st.container(border=True):
        st.markdown(render_preview_html(df, hv, F), unsafe_allow_html=True)

    st.markdown("#### Xuất báo giá")
    cust = st.checkbox("Bản gửi khách (không có ghi chú nội bộ, không tô vàng)", key="cust_mode")
    base = st.text_input("Tên file (không kèm đuôi)",
                         value=f"BaoGia_TSC_{datetime.now().strftime('%Y%m%d_%H%M')}", key="out_base")
    b1, b2 = st.columns(2)
    if b1.button("📊 Xuất Excel", type="primary", use_container_width=True, key="btn_xl"):
        xl = build_excel(df, H, F, customer=cust)
        st.session_state["out_xl"] = (xl, base)
    if b2.button("📄 Xuất PDF", type="primary", use_container_width=True, key="btn_pdf"):
        xl = build_excel(df, H, F, customer=cust)  # PDF và Excel cùng một nguồn, nội dung giống nhau
        pdf = xlsx_to_pdf(xl)
        if pdf is None:
            st.warning("Chưa có LibreOffice nên PDF là bản đơn giản.")
            pdf = build_quote_pdf(df, hv, F, customer=cust)
        st.session_state["out_pdf"] = (pdf, base, xl)

    # Lưu lịch sử khi đã có file
    tong = quote_totals(df)[2]
    so_dong = len(df)
    if st.session_state.get("out_xl") or st.session_state.get("out_pdf"):
        xl_data = st.session_state.get("out_xl", (None, None))[0]
        pdf_data = st.session_state.get("out_pdf", (None, None, None))[0]
        ten = (st.session_state.get("out_xl") or st.session_state.get("out_pdf"))[1]
        if xl_data is None and pdf_data is not None:
            xl_data = st.session_state.get("out_pdf")[2]
        if st.session_state.get("saved_key") != (ten, tong, so_dong):
            lich_su.save_quote({"so": H["no"], "khach": H["to"], "du_an": H["project"], "tong": round(tong),
                                "so_dong": so_dong, "ban_gui_khach": bool(cust), "ten_file": ten},
                               xlsx_bytes=xl_data, pdf_bytes=pdf_data)
            st.session_state.saved_key = (ten, tong, so_dong)
            st.success("Đã lưu vào Lịch sử báo giá.")
        if st.session_state.get("out_xl"):
            st.download_button("⬇️ Tải file Excel", data=st.session_state["out_xl"][0],
                               file_name=f"{base}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_xl")
        if st.session_state.get("out_pdf"):
            st.download_button("⬇️ Tải file PDF", data=st.session_state["out_pdf"][0],
                               file_name=f"{base}.pdf", mime="application/pdf", key="dl_pdf")


# =============================== LỊCH SỬ ===============================
def page_history():
    kw = st.text_input("Tìm theo số báo giá, khách, dự án", key="hist_kw")
    rows = lich_su.list_quotes(kw)
    if not rows:
        st.info("Chưa có báo giá nào được xuất.")
        return
    table = pd.DataFrame([{"Ngày": r.get("ngay"), "Số báo giá": r.get("so"), "Khách hàng": r.get("khach"),
                           "Dự án": r.get("du_an"), "Tổng tiền (VND)": f"{r.get('tong', 0):,.0f}",
                           "Số dòng": r.get("so_dong"), "Gửi khách": "Có" if r.get("ban_gui_khach") else "Không"}
                          for r in rows])
    st.dataframe(table, use_container_width=True, hide_index=True)
    labels = [f"{r.get('ngay')} · {r.get('so')} · {r.get('khach')}" for r in rows]
    pick = st.selectbox("Chọn báo giá để tải lại", range(len(rows)), format_func=lambda i: labels[i],
                        key="hist_pick")
    r = rows[pick]
    d1, d2 = st.columns(2)
    xl = lich_su.read_file(r.get("file_xlsx"))

    pdf = lich_su.read_file(r.get("file_pdf"))
    if xl:
        d1.download_button("⬇️ Tải Excel", xl, file_name=f"{r.get('ten_file') or 'bao_gia'}.xlsx", key="h_xl")
    if pdf:
        d2.download_button("⬇️ Tải PDF", pdf, file_name=f"{r.get('ten_file') or 'bao_gia'}.pdf", key="h_pdf")
    if not xl and not pdf:
        st.caption("Báo giá này không còn file đính kèm.")


# =============================== DANH MỤC ===============================
def page_catalog():
    if not os.path.exists(PRICE_PATH):
        st.info("Không tìm thấy file bảng giá.")
        return
    sheet_names = [s for s in cached_price_sheet_names(PRICE_PATH, os.path.getmtime(PRICE_PATH))
                   if s.upper() != PLACEHOLDER_SHEET]
    all_items = load_price_items(PRICE_PATH, sheet_names)
    if not all_items:
        st.info("Không đọc được bảng giá.")
        return

    # Đếm số mục của từng hãng (sheet) để hiện ngay trong danh sách chọn
    count = {}
    for it in all_items:
        count[it["sheet"]] = count.get(it["sheet"], 0) + 1
    ALL = "— Tất cả hãng —"
    options = [ALL] + [f"{sn}  ({count.get(sn, 0)} mục)" for sn in sheet_names if count.get(sn)]
    c1, c2 = st.columns([2, 3])
    choice = c1.selectbox("Chọn hãng (sheet bảng giá)", options, key="cat_brand")
    kw = c2.text_input("Tìm trong hãng này (mã, tên)", key="cat_kw").strip().lower()

    brand = None if choice == ALL else choice.rsplit("  (", 1)[0]
    base = all_items if brand is None else [it for it in all_items if it["sheet"] == brand]
    shown = [it for it in base if not kw or kw in f"{it['ma']} {it['ten']}".lower()]
    st.caption(f"Hiển thị {len(shown)} / {len(base)} mục. Sửa cột **Đơn giá** rồi bấm **Lưu đơn giá**.")

    table = pd.DataFrame([{"Hãng": it["sheet"], "Mã": it["ma"], "Tên": it["ten"], "Xuất xứ": it["xuat_xu"],
                           "ĐVT": it["dvt"], "Đơn giá (VND)": it["don_gia"]} for it in shown])
    edited = st.data_editor(
        table, key=f"cat_editor_{choice}_{kw}", use_container_width=True, height=480, hide_index=True,
        disabled=["Hãng", "Mã", "Tên", "Xuất xứ", "ĐVT"],
        column_config={"Đơn giá (VND)": st.column_config.NumberColumn(
            "Đơn giá (VND)", format="%,.0f", min_value=0)})
    if st.button("💾 Lưu đơn giá", type="primary", key="cat_save"):
        changed = 0
        for i, it in enumerate(shown):
            new_price = float(edited.iloc[i]["Đơn giá (VND)"] or 0)
            if abs(new_price - float(it["don_gia"] or 0)) > 0.0001:
                save_override(it, new_price)
                changed += 1
        if changed:
            cached_price_items.clear()
            st.success(f"Đã lưu {changed} đơn giá. Các báo giá mới sẽ dùng giá này.")
            st.rerun()
        else:
            st.info("Không có đơn giá nào thay đổi.")


if page == "📝 Báo giá mới":
    page_quote()
elif page == "🗂️ Lịch sử báo giá":
    page_history()
else:
    page_catalog()
