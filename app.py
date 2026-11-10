import os
import copy
from datetime import datetime
from io import BytesIO

import openpyxl
import pandas as pd
import streamlit as st

from ai_assistant import (
    COLUMNS, KnowledgeStore, audit_quote, build_quote_rows, clean,
    detect_boq_columns, has_llm, sync_codes,
    load_price_items, promote_header, suggest_matches, apply_approved,
)

# --- ĐƯỜNG DẪN (luôn tính theo thư mục chứa app.py) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "MauBaoGia.xlsx")
PRICE_LIST_PATH = os.path.join(BASE_DIR, "BangGia.xlsx")
KNOWLEDGE_PATH = os.path.join(BASE_DIR, "learned_knowledge.json")
PLACEHOLDER_SHEET = "XXXXXXXXX"
QUOTE_SHEET = "BÁO GIÁ"

# --- CẤU HÌNH TRANG & GIAO DIỆN ---
st.set_page_config(
    page_title="QN TECH - Hệ Thống Báo Giá",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #121212; color: #E0E0E0; }
    [data-testid="stSidebar"] { background-color: #1A1A24; border-right: 1px solid #333; }
    h1 { color: #E85D04 !important; font-family: 'Segoe UI', Tahoma, sans-serif; font-weight: 700 !important; border-bottom: 2px solid #0C2340; padding-bottom: 10px; }
    h2, h3 { color: #4A90E2 !important; font-family: 'Segoe UI', Tahoma, sans-serif; }
    .stButton > button[data-baseweb="button"] { background-color: #E85D04; color: white; border-radius: 5px; border: none; font-weight: bold; transition: all 0.3s ease; }
    .stButton > button[data-baseweb="button"]:hover { background-color: #FF7A00; box-shadow: 0px 4px 10px rgba(232, 93, 4, 0.4); }
    .stAlert { background-color: #1E1E2E; border: 1px solid #333; color: #E0E0E0; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; border: 1px solid #333; }
    hr { border-top: 1px solid #333; }
    .streamlit-expanderHeader { background-color: #1A1A24; color: #4A90E2; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
    <h1 style="margin: 0; border: none; font-size: 36px;">
        <span style="color: #4A90E2;">💠 QN TECH</span> <span style="color: #E0E0E0; font-size: 24px; font-weight: 400;">| Hệ Thống Xử Lý Báo Giá Thông Minh</span>
    </h1>
</div>
""", unsafe_allow_html=True)

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
    df.at[idx, "Mã hàng"] = it["ma"]
    df.at[idx, "Diễn giải"] = it["ten"]
    df.at[idx, "Xuất xứ"] = it["xuat_xu"]
    df.at[idx, "ĐVT"] = it["dvt"] or df.at[idx, "ĐVT"]
    df.at[idx, "Đơn giá (VND)"] = it["don_gia"]
    df.at[idx, "Nguồn"] = "Chọn tay"
    df.at[idx, "Duyệt"] = True


# --- SIDEBAR: TẢI DỮ LIỆU ---
st.sidebar.markdown("<h2 style='text-align: center; color: #E85D04;'>📥 TẢI DỮ LIỆU</h2>", unsafe_allow_html=True)
uploaded_boq = st.sidebar.file_uploader(
    "Tải lên File BOQ khối lượng gốc (Excel/CSV)", type=["xlsx", "xls", "csv"])

boq_sheet_selected = None
if uploaded_boq is not None and uploaded_boq.name.lower().endswith(("xlsx", "xls")):
    try:
        uploaded_boq.seek(0)
        xl_boq = pd.ExcelFile(uploaded_boq)
        st.sidebar.markdown("---")
        boq_sheet_selected = st.sidebar.selectbox("📂 Chọn Sheet chứa khối lượng (BOQ):", xl_boq.sheet_names)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file BOQ: {e}")

price_items = []
if os.path.exists(PRICE_LIST_PATH):
    try:
        all_sheets = pd.ExcelFile(PRICE_LIST_PATH).sheet_names
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
st.sidebar.markdown("<h3 style='color: #4A90E2;'>🧠 TRỢ LÝ AI</h3>", unsafe_allow_html=True)
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
        uploaded_boq.seek(0)
        if uploaded_boq.name.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded_boq, header=None)
        else:
            raw = pd.read_excel(uploaded_boq, sheet_name=boq_sheet_selected, header=None)
    except Exception as e:
        st.error(f"Lỗi đọc file BOQ: {e}")

    if raw is not None:
        raw = raw.dropna(how="all")
        if raw.empty:
            st.error(f"❌ Sheet '{boq_sheet_selected}' không có dữ liệu.")
        else:
            df_raw, header_row = promote_header(raw)
            detected = detect_boq_columns(df_raw)
            df_raw = df_raw.dropna(how="all")
            all_cols = [str(c) for c in df_raw.columns]
            file_tag = f"{uploaded_boq.name}_{boq_sheet_selected}"
            labels = {"name": "Tên hàng hóa", "qty": "Số lượng (SL)",
                      "unit": "Đơn vị", "spec": "Thông số kỹ thuật"}
            opts = ["(không có)"] + all_cols
            auto_ok = bool(detected.get("name") and detected.get("qty"))
            with st.expander("🔎 Xác nhận cột (hệ thống đã tự chọn, bạn có thể đổi)", expanded=not auto_ok):
                st.caption(f"Tiêu đề lấy từ dòng {header_row + 1} của sheet. "
                           "Nếu cột nào sai, chọn lại trong danh sách.")
                c1, c2 = st.columns(2)
                cols = {}
                for i, (key, label) in enumerate(labels.items()):
                    default = detected.get(key)
                    idx = opts.index(default) if default in opts else 0
                    chosen = (c1 if i % 2 == 0 else c2).selectbox(
                        label, opts, index=idx, key=f"col_{key}_{file_tag}")
                    cols[key] = None if chosen == "(không có)" else chosen
                cols["price"] = detected.get("price")
                st.caption("Xem trước 8 dòng đầu của sheet:")
                st.dataframe(raw.head(8), use_container_width=True)

            if not cols["name"] or not cols["qty"]:
                st.error("❌ Chưa chọn đủ cột Tên hàng và Số lượng. Mở mục 'Xác nhận cột' để chọn.")
            elif st.sidebar.button("🚀 BẮT ĐẦU XỬ LÝ BOQ", type="primary", use_container_width=True):
                if not price_items:
                    st.warning("Chưa có bảng giá nào được chọn ở sidebar. Các dòng sẽ không được ghép mã.")
                with st.spinner("AI đang phân tích và gợi ý mã hàng..."):
                    rows, stats = build_quote_rows(df_raw, cols, price_items, store, use_llm=use_llm)
                    st.session_state.df_editable = pd.DataFrame(rows, columns=COLUMNS)
                    st.session_state.learned_keys = set()
                    st.session_state.last_stats = stats
                    st.session_state.editor_ver += 1

# --- GIAO DIỆN CHỈNH SỬA ---
if st.session_state.df_editable is not None:
    df = recalc(st.session_state.df_editable)
    st.markdown("<hr style='border: 1px solid #4A90E2;'>", unsafe_allow_html=True)
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
                        store.log_correction(query, before, it["ma"])
                    st.session_state.df_editable = recalc(df)
                    st.session_state.editor_ver += 1
                    st.rerun()

    # --- Xóa dòng đã tick ---
    col_act1, col_act2, _ = st.columns([1, 1, 3])
    with col_act1:
        if st.button("🗑️ XÓA DÒNG ĐÃ TICK", use_container_width=True):
            st.session_state.df_editable = df[df["Xóa"] != True].reset_index(drop=True)
            st.session_state.editor_ver += 1
            st.rerun()
    with col_act2:
        if st.button("✅ DUYỆT TẤT CẢ ĐỀ XUẤT", use_container_width=True):
            df.loc[df["Mã đề xuất"] != "", "Duyệt"] = True
            st.session_state.df_editable = df
            st.session_state.editor_ver += 1
            st.rerun()

    # --- Bảng chỉnh sửa ---
    code_options = sorted({p["ma"] for p in price_items if p["ma"]})
    editor_key = f"boq_editor_{st.session_state.editor_ver}"
    edited_df = st.data_editor(
        df,
        key=editor_key,
        num_rows="dynamic",
        use_container_width=True,
        height=500,
        column_config={
            "Xóa": st.column_config.CheckboxColumn("🗑️ Xóa", help="Tick để xóa dòng này", default=False),
            "Duyệt": st.column_config.CheckboxColumn("✅ Duyệt", help="Tick để điền mã đề xuất vào file", default=False),
            "BOQ gốc": st.column_config.TextColumn("BOQ gốc", disabled=True, width="large"),
            "Mã đề xuất": st.column_config.TextColumn("Mã đề xuất (AI)", disabled=True),
            "Độ khớp": st.column_config.NumberColumn("Độ khớp", format="%.2f", disabled=True),
            "Nguồn": st.column_config.TextColumn("Nguồn", disabled=True),
            "Mã hàng": st.column_config.SelectboxColumn(
                "Mã hàng", options=code_options, help="Chọn mã từ bảng giá. Tên và giá sẽ tự điền."),
            "Gợi ý": st.column_config.TextColumn("Gợi ý khác", disabled=True),
            "SL": st.column_config.NumberColumn("SL", format="%.2f", min_value=0),
            "Đơn giá (VND)": st.column_config.NumberColumn("Đơn giá (VND)", format="%,.0f", min_value=0),
            "Thành tiền (VND)": st.column_config.NumberColumn("Thành tiền (VND)", format="%,.0f", disabled=True),
        },
    )

    # Đồng bộ mã người dùng chọn, rồi điền các dòng đã duyệt
    editor_state = st.session_state.get(editor_key, {})
    edited_df = recalc(edited_df)
    edited_rows = editor_state.get("edited_rows", {}) if isinstance(editor_state, dict) else {}
    synced = sync_codes(edited_df, edited_rows, price_items, store, st.session_state.learned_keys)
    applied_now = apply_approved(edited_df, price_items, store, st.session_state.learned_keys)
    st.session_state.df_editable = edited_df
    if synced or applied_now:
        st.session_state.editor_ver += 1
        st.toast(f"✅ Đã cập nhật {synced + applied_now} dòng và ghi nhớ cho lần sau")
        st.rerun()
    df = st.session_state.df_editable

    # --- Kiểm tra trước khi xuất ---
    audit_rows = [{"ma": clean(r["Mã hàng"]), "sl": r["SL"], "don_gia": r["Đơn giá (VND)"]}
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

    # --- Xuất Excel ---
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

        def shift_rows(at_row, amount):
            """Thêm (amount > 0) hoặc xóa (amount < 0) dòng tại at_row.
            Giữ nguyên merge ô và chiều cao dòng."""
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
                    continue  # merge nằm trong vùng đã xóa
                ws.merge_cells(start_row=r1 + amount, start_column=c1,
                               end_row=r2 + amount, end_column=c2)

        # Header
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

        # Tìm dòng "Cộng/Tổng" làm ranh giới vùng hàng hóa
        start_row = 14
        end_row = start_row
        for r in range(start_row, start_row + 150):
            c2 = str(ws.cell(row=r, column=2).value or "").lower()
            c3 = str(ws.cell(row=r, column=3).value or "").lower()
            if "cộng" in c2 or "tổng" in c2 or "cộng" in c3 or "tổng" in c3:
                end_row = r
                break
        available = end_row - start_row
        n = len(data)

        if n > available:
            extra = n - available
            shift_rows(end_row, extra)
            for r in range(end_row, end_row + extra):
                for col in range(1, 10):
                    s_cell = ws.cell(row=end_row - 1, column=col)
                    t_cell = ws.cell(row=r, column=col)
                    if s_cell.has_style:
                        t_cell.font = copy.copy(s_cell.font)
                        t_cell.border = copy.copy(s_cell.border)
                        t_cell.fill = copy.copy(s_cell.fill)
                        t_cell.number_format = s_cell.number_format
                        t_cell.alignment = copy.copy(s_cell.alignment)
            end_row += extra
        elif n < available:
            fewer = available - n
            shift_rows(start_row + n, -fewer)
            end_row -= fewer

        for i, (_, r) in enumerate(data.iterrows()):
            rr = start_row + i
            safe_write(rr, 1, i + 1)
            safe_write(rr, 2, clean(r["Mã hàng"]))
            safe_write(rr, 3, clean(r["Diễn giải"]))
            safe_write(rr, 4, clean(r["Xuất xứ"]))
            safe_write(rr, 5, clean(r["ĐVT"]) or "Cái")
            safe_write(rr, 6, float(r["SL"]))
            safe_write(rr, 7, float(r["Đơn giá (VND)"]))
            safe_write(rr, 8, f"=F{rr}*G{rr}")
            safe_write(rr, 9, clean(r["Ghi chú"]))

        safe_write(end_row, 8, f"=SUM(H{start_row}:H{end_row - 1})")

        for r in range(end_row, ws.max_row + 1):
            c1 = str(ws.cell(row=r, column=1).value or "").strip().lower()
            c2 = str(ws.cell(row=r, column=2).value or "").strip().lower()
            c3 = str(ws.cell(row=r, column=3).value or "").strip().lower()
            if "lưu ý" in c2:
                safe_write(r, 3, f_note)
            elif "giao hàng" in c2:
                safe_write(r, 3, f_deliv1)
            elif "nhập hàng" in c3 or "giấy lưu thông" in c3:
                safe_write(r, 3, f_deliv2)
            elif "thanh toán" in c2:
                safe_write(r, 3, f_payment)

        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    output_filename = st.text_input(
        "Tên file Excel xuất ra:", value=f"BaoGia_QNTech_{datetime.now().strftime('%Y%m%d')}.xlsx", key="out_name")

    if st.button("📥 XUẤT FILE BÁO GIÁ", use_container_width=True):
        if n_err > 0:
            st.warning(f"Báo giá còn {n_err} lỗi (đơn giá = 0). Bạn vẫn có thể xuất, nhưng nên kiểm tra lại.")
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
