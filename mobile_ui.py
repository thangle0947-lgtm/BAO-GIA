"""
mobile_ui.py - Giao diện tối ưu cho ĐIỆN THOẠI (một cột, mỗi hàng hóa là một thẻ, nút to).

Được app.py gọi khi người dùng chọn "ĐIỆN THOẠI". Dùng chung dữ liệu, bảng giá, bộ nhớ học
và hàm xuất Excel/PDF với giao diện máy tính, nên báo giá ra giống hệt nhau.
4 bước: 1 BOQ -> 2 Hàng hóa -> 3 Thông tin khách -> 4 Xuất file.
"""
import math
import os
from datetime import datetime

import pandas as pd
import streamlit as st

from ai_assistant import (
    COLUMNS, apply_approved, audit_quote, build_quote_rows, clean, code_of, detect_boq_columns,
    fill_item, find_header_row, find_layout, has_llm, item_key, layout_columns, promote_header,
)

PAGE_SIZE = 8
STEPS = ["📥 1 BOQ", "📋 2 Hàng", "👤 3 Khách", "📤 4 Xuất"]
NONE_LABEL = "(chưa chọn mã)"
FILTERS = ["Tất cả", "Chưa có mã", "Chưa có giá"]

MOBILE_CSS = """
<style>
    [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"] { display: none; }
    .block-container { padding: 0.8rem 0.7rem 4rem 0.7rem !important; max-width: 680px; margin: 0 auto; }
    .brand-card { padding: 8px 12px !important; border-radius: 12px !important; margin-bottom: 6px !important; }
    .brand-sub { font-size: 14px !important; margin: 2px 0 4px !important; }
    .gold-bar { margin-bottom: 8px !important; }
    .stButton > button, .stDownloadButton > button { min-height: 48px; font-size: 16px; }
    input, textarea, [data-baseweb="select"] * { font-size: 16px !important; }
    h3 { font-size: 20px !important; }
    [data-testid="stRadio"] label { font-size: 14px; }
</style>
"""


# --------------------------------------------------------------- trạng thái
def _init_state():
    defaults = {"m_ver": 0, "m_page": 0, "m_filter": FILTERS[0], "m_step": STEPS[0], "m_goto": None}
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v
    if "m_info" not in st.session_state:
        st.session_state.m_info = {
            "to": "", "add": "", "att": "", "mobile": "", "email": "", "ref": "",
            "no": f"BG-{datetime.now().strftime('%d%m%y')}-01", "date": datetime.now().date(),
            "revise": "0", "validity": "30 ngày", "project": "",
            "note": "Báo giá đã bao gồm vận chuyển và giấy thông hành",
            "deliv1": "Tại địa chỉ dự án bên mua.",
            "deliv2": "Nhập hàng 04-06 Tuần + giấy lưu thông 04-06 tuần",
            "bank": "",
            "payment": "Tạm ứng 30% khi đặt hàng\nThanh toán 70% còn lại trước khi giao hàng",
        }


def _money(v):
    return f"{v:,.0f}" if v else "0"


def _goto(i):
    st.session_state.m_goto = i


# ------------------------------------------------------------------- callback
def _cb_num(idx, col, widget_key):
    df = st.session_state.df_editable
    if df is not None and idx in df.index:
        try:
            df.at[idx, col] = float(st.session_state.get(widget_key) or 0)
        except (TypeError, ValueError):
            pass


def _cb_code(c, idx, widget_key, price_key, by_key):
    choice = st.session_state.get(widget_key)
    df = st.session_state.df_editable
    item = by_key.get(choice)
    if df is None or idx not in df.index or item is None:
        return
    df = c.recalc(df)
    fill_item(df, idx, item)
    df.at[idx, "Nguồn"] = "Chọn tay"
    df.at[idx, "Duyệt"] = True
    boq = clean(df.at[idx, "BOQ gốc"])
    if boq:
        c.store.learn_mapping(boq, item)
        c.store.log_correction(boq, "", item["ma"])
    df, merged = c.merge_and_remap(df)           # trùng mã + mô tả thì gộp, cộng SL
    st.session_state.df_editable = c.recalc(df)
    if merged:
        st.session_state.m_ver += 1               # chỉ số dòng đã đổi: dựng lại các thẻ
    else:
        st.session_state[price_key] = float(item["don_gia"])  # cập nhật ô giá của thẻ này


def _cb_accept(idx):
    df = st.session_state.df_editable
    if df is not None and idx in df.index:
        df.at[idx, "Duyệt"] = True
        st.session_state.m_ver += 1


def _shift_seen_after_delete(idx):
    new = set()
    for k in st.session_state.learned_keys:
        if isinstance(k, tuple) and k and k[0] == "ma":
            if k[1] == idx:
                continue
            new.add(("ma", k[1] - 1 if k[1] > idx else k[1], k[2]))
        elif isinstance(k, tuple) and k and isinstance(k[0], int):
            if k[0] == idx:
                continue
            new.add((k[0] - 1 if k[0] > idx else k[0],) + tuple(k[1:]))
        else:
            new.add(k)
    st.session_state.learned_keys = new


def _cb_delete(idx):
    df = st.session_state.df_editable
    if df is not None and idx in df.index:
        _shift_seen_after_delete(idx)
        st.session_state.df_editable = df.drop(index=idx).reset_index(drop=True)
        st.session_state.m_ver += 1


def _cb_page(delta):
    st.session_state.m_page = max(0, st.session_state.m_page + delta)


def _cb_filter():
    st.session_state.m_page = 0


def _cb_add_row(c):
    df = c.recalc(st.session_state.df_editable)
    row = {col: "" for col in COLUMNS}
    row.update({"Xóa": False, "Duyệt": False, "SL": 1.0, "Đơn giá (VND)": 0.0, "Thành tiền (VND)": 0.0,
                "Độ khớp": 0.0, "ĐVT": "Cái", "Gợi ý": "Dòng thêm tay"})
    df = pd.concat([df, pd.DataFrame([row], columns=COLUMNS)], ignore_index=True)
    st.session_state.df_editable = df
    st.session_state.m_filter = FILTERS[0]
    st.session_state.m_page = 10 ** 6            # nhảy tới trang cuối
    st.session_state.m_ver += 1


def _cb_approve_all(c):
    df = c.recalc(st.session_state.df_editable)
    df.loc[df["Mã đề xuất"] != "", "Duyệt"] = True
    st.session_state.df_editable = df
    st.session_state.m_ver += 1


def _cb_merge(c):
    df, merged = c.merge_and_remap(c.recalc(st.session_state.df_editable))
    st.session_state.df_editable = c.recalc(df)
    st.session_state.m_ver += 1
    st.session_state.m_merge_note = merged


# --------------------------------------------------------------- bước 1: BOQ
def _step_boq(c, price_items, all_price_sheets):
    st.markdown("### 📥 Tải file BOQ")
    up = st.file_uploader("Chọn file BOQ (Excel/CSV)", type=["xlsx", "xls", "csv"], key="m_boq")

    if all_price_sheets:
        with st.expander(f"📋 Bảng giá áp dụng ({len(price_items)} mục)"):
            current = st.session_state.get("m_price_sel", all_price_sheets)
            st.multiselect("Sheet bảng giá:", all_price_sheets, default=current, key="m_ps_w",
                           on_change=lambda: st.session_state.update(m_price_sel=st.session_state.m_ps_w))
    elif not os.path.exists(c.PRICE_LIST_PATH):
        st.warning("Chưa thấy file BangGia.xlsx cạnh app.py.")
    st.caption(f"🧠 Đã học: {len(c.store.data['mappings'])} ghép mã")

    if up is None:
        st.info("Chọn file BOQ khối lượng để bắt đầu.")
        if st.session_state.get("df_editable") is not None:
            st.button("➡️ Về danh sách hàng đang làm", on_click=_goto, args=(1,), use_container_width=True)
        return

    sheet = None
    if up.name.lower().endswith(("xlsx", "xls")):
        try:
            names = c.cached_boq_sheet_names(up.getvalue())
            sheet = st.selectbox("📂 Sheet chứa khối lượng:", names, key="m_boq_sheet")
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi đọc file BOQ: {e}")
            return
    try:
        raw = c.cached_boq_raw(up.getvalue(), up.name, sheet).copy()
    except Exception as e:  # noqa: BLE001
        st.error(f"Lỗi đọc file BOQ: {e}")
        return
    raw = raw.dropna(how="all").reset_index(drop=True)
    if raw.empty:
        st.error("Sheet này không có dữ liệu.")
        return

    file_tag = f"{up.name}_{sheet}"
    layout, layout_row = find_layout(raw, c.store.data["layouts"])
    default_header = layout_row if layout else find_header_row(raw)
    labels = {"name": "Cột TÊN HÀNG", "qty": "Cột SỐ LƯỢNG", "unit": "Cột ĐƠN VỊ", "spec": "Cột THÔNG SỐ / MÔ TẢ"}

    with st.expander("🔎 Cách đọc file (kiểm tra cột)", expanded=layout is None):
        if layout:
            st.success("Đã nhận ra cấu trúc file, cột đã được nhớ từ lần trước.")
        else:
            st.info("File mới: hãy chọn đúng cột. Hệ thống sẽ nhớ cho lần sau.")
        header_ui = st.number_input("Dòng tiêu đề (số dòng trong Excel)", min_value=1,
                                    max_value=max(1, min(30, len(raw))), value=int(default_header) + 1,
                                    step=1, key=f"m_hdr_{file_tag}")
        header_row = int(header_ui) - 1
        df_raw, _ = promote_header(raw, header_row)
        df_raw = df_raw.dropna(how="all")
        all_cols = [str(x) for x in df_raw.columns]
        heur = detect_boq_columns(df_raw)
        from_layout = layout_columns(df_raw, layout) if (layout and header_row == layout_row) else {}
        guess = {k: (from_layout.get(k) or heur.get(k)) for k in ["name", "qty", "unit", "spec", "price"]}
        opts = ["(không có)"] + all_cols
        cols = {}
        for key in ["name", "qty", "unit", "spec"]:
            default = guess.get(key)
            idx = opts.index(default) if default in opts else 0
            chosen = st.selectbox(labels[key], opts, index=idx, key=f"m_col_{key}_{file_tag}_{header_row}")
            cols[key] = None if chosen == "(không có)" else chosen
        cols["price"] = guess.get("price")
        st.caption("8 dòng đầu của sheet:")
        st.dataframe(raw.head(8), use_container_width=True, height=220)

    if not cols["name"] or not cols["qty"]:
        st.error("Chưa chọn đủ cột Tên hàng và Số lượng. Mở mục 'Cách đọc file' để chọn.")
        return

    use_llm = False
    if has_llm():
        use_llm = st.checkbox("Dùng Claude cho dòng khó", value=False, key="m_llm")

    if st.button("🚀 XỬ LÝ BOQ", type="primary", use_container_width=True, key="m_process"):
        c.store.save_layout(raw, header_row, list(df_raw.columns), cols)
        if not price_items:
            st.warning("Chưa có bảng giá, các dòng sẽ không được ghép mã.")
        with st.spinner("AI đang phân tích và gợi ý mã hàng..."):
            rows, stats = build_quote_rows(df_raw, cols, price_items, c.store, use_llm=use_llm)
        st.session_state.df_editable = pd.DataFrame(rows, columns=COLUMNS)
        st.session_state.learned_keys = set()
        st.session_state.last_stats = stats
        st.session_state.m_ver += 1
        st.session_state.m_page = 0
        st.session_state.m_filter = FILTERS[0]
        st.session_state.m_goto = 1
        st.rerun()


# ----------------------------------------------------------- bước 2: hàng hóa
def _card(c, df, idx, code_pos, code_opts, by_key, ver):
    row = df.loc[idx]
    code = clean(row["Mã hàng"])
    price = float(row["Đơn giá (VND)"] or 0)
    qty = float(row["SL"] or 0)
    missing = (not code) or price <= 0
    with st.container(border=True):
        text = clean(row["BOQ gốc"]) or clean(row["Diễn giải"])
        st.markdown(f"**{'⚠️' if missing else '✅'} #{int(idx) + 1}** · {text}")

        prop = clean(row["Mã đề xuất"])
        if prop and not code:
            st.caption(f"💡 Đề xuất (khớp {float(row['Độ khớp'] or 0):.2f}): {prop}")
            st.button("✅ Dùng mã đề xuất", key=f"m_acc_{ver}_{idx}", on_click=_cb_accept, args=(idx,),
                      use_container_width=True)
        elif not code and clean(row["Gợi ý"]):
            st.caption(clean(row["Gợi ý"]))

        ck, pk, qk = f"m_code_{ver}_{idx}", f"m_pr_{ver}_{idx}", f"m_sl_{ver}_{idx}"
        if ck not in st.session_state:
            st.session_state[ck] = code if code in by_key else NONE_LABEL
        if pk not in st.session_state:
            st.session_state[pk] = price
        if qk not in st.session_state:
            st.session_state[qk] = qty
        st.selectbox("Mã hàng (gõ để tìm)", [NONE_LABEL] + code_opts, key=ck,
                     on_change=_cb_code, args=(c, idx, ck, pk, by_key))
        if code:
            st.caption(f"{clean(row['Diễn giải'])} · {clean(row['Xuất xứ'])} · {clean(row['ĐVT'])}")

        c1, c2 = st.columns(2)
        with c1:
            st.number_input("Số lượng", min_value=0.0, step=1.0, format="%g", key=qk,
                            on_change=_cb_num, args=(idx, "SL", qk))
        with c2:
            st.number_input("Đơn giá (VND)", min_value=0.0, step=1000.0, format="%.0f", key=pk,
                            on_change=_cb_num, args=(idx, "Đơn giá (VND)", pk))
        total = float(df.at[idx, "SL"] or 0) * float(df.at[idx, "Đơn giá (VND)"] or 0)
        c3, c4 = st.columns([3, 1])
        with c3:
            st.markdown(f"Thành tiền: **{_money(total)} đ**")
        with c4:
            st.button("🗑️", key=f"m_del_{ver}_{idx}", on_click=_cb_delete, args=(idx,),
                      help="Xóa dòng này")


def _step_items(c, price_items):
    df = st.session_state.get("df_editable")
    if df is None:
        st.info("Chưa có dữ liệu. Quay lại bước 1 để tải và xử lý BOQ.")
        st.button("⬅️ Về bước 1", on_click=_goto, args=(0,), use_container_width=True)
        return

    # Điền mã cho các dòng đã duyệt, rồi gộp dòng trùng mã + mô tả (giống giao diện máy tính)
    df = c.recalc(df)
    if apply_approved(df, price_items, c.store, st.session_state.learned_keys):
        df, _ = c.merge_and_remap(df)
        st.session_state.df_editable = c.recalc(df)
        st.session_state.m_ver += 1
        st.rerun()
    df = c.recalc(st.session_state.df_editable)
    ver = st.session_state.m_ver

    by_key = {item_key(p): p for p in price_items if p["ma"]}
    code_opts = sorted(by_key)
    code_pos = {k: i for i, k in enumerate(code_opts)}

    n_nocode = int(sum(1 for _, r in df.iterrows() if not code_of(r["Mã hàng"])))
    n_noprice = int(sum(1 for _, r in df.iterrows() if float(r["Đơn giá (VND)"] or 0) <= 0))
    cong, thue, tong = c.quote_totals(df)
    st.markdown("### 📋 Hàng hóa")
    st.caption(f"{len(df)} dòng · ⚠️ {n_nocode} chưa có mã · {n_noprice} chưa có giá")
    st.markdown(f"Tạm tính: **{_money(cong)} đ** · Tổng sau thuế: **{_money(tong)} đ**")

    if st.session_state.get("m_merge_note") is not None:
        n = st.session_state.pop("m_merge_note")
        st.success(f"🔗 Đã gộp {n} dòng trùng." if n else "Không có dòng nào trùng để gộp.")
    if st.session_state.merge_msg:
        st.success(f"🔗 Đã gộp {st.session_state.merge_msg} dòng trùng mã + mô tả, số lượng đã cộng dồn.")
        st.session_state.merge_msg = 0

    b1, b2 = st.columns(2)
    with b1:
        st.button("✅ Duyệt tất cả đề xuất", on_click=_cb_approve_all, args=(c,), use_container_width=True)
    with b2:
        st.button("🔗 Gộp dòng trùng", on_click=_cb_merge, args=(c,), use_container_width=True)

    st.radio("Hiện", FILTERS, horizontal=True, key="m_filter", on_change=_cb_filter)

    if st.session_state.m_filter == "Chưa có mã":
        idxs = [i for i in df.index if not code_of(df.at[i, "Mã hàng"])]
    elif st.session_state.m_filter == "Chưa có giá":
        idxs = [i for i in df.index if float(df.at[i, "Đơn giá (VND)"] or 0) <= 0]
    else:
        idxs = list(df.index)

    n_pages = max(1, math.ceil(len(idxs) / PAGE_SIZE))
    page = min(st.session_state.m_page, n_pages - 1)
    st.session_state.m_page = page

    def nav(suffix):
        n1, n2, n3 = st.columns([1, 2, 1])
        with n1:
            st.button("◀", key=f"m_prev_{suffix}", on_click=_cb_page, args=(-1,), disabled=page <= 0,
                      use_container_width=True)
        with n2:
            st.markdown(f"<div style='text-align:center;padding-top:10px'>Trang {page + 1}/{n_pages}</div>",
                        unsafe_allow_html=True)
        with n3:
            st.button("▶", key=f"m_next_{suffix}", on_click=_cb_page, args=(1,), disabled=page >= n_pages - 1,
                      use_container_width=True)

    nav("top")
    if not idxs:
        st.info("Không có dòng nào theo bộ lọc này.")
    for idx in idxs[page * PAGE_SIZE:(page + 1) * PAGE_SIZE]:
        _card(c, df, idx, code_pos, code_opts, by_key, ver)
    nav("bottom")

    st.button("➕ Thêm dòng trống", on_click=_cb_add_row, args=(c,), use_container_width=True)
    st.button("➡️ Tiếp: Thông tin khách hàng", type="primary", on_click=_goto, args=(2,),
              use_container_width=True)


# ------------------------------------------------------------ bước 3: khách hàng
def _step_info():
    info = st.session_state.m_info
    st.markdown("### 👤 Thông tin khách hàng")
    info["to"] = st.text_input("To (Tên công ty)", value=info["to"], key="mw_to")
    info["add"] = st.text_input("Add (Địa chỉ)", value=info["add"], key="mw_add")
    info["att"] = st.text_input("Att (Người liên hệ)", value=info["att"], key="mw_att")
    info["mobile"] = st.text_input("Mobile (Số điện thoại)", value=info["mobile"], key="mw_mobile")
    info["email"] = st.text_input("Email", value=info["email"], key="mw_email")
    info["ref"] = st.text_input("Ref", value=info["ref"], key="mw_ref")
    info["project"] = st.text_input("Project (Tên dự án)", value=info["project"], key="mw_project")
    c1, c2 = st.columns(2)
    with c1:
        info["no"] = st.text_input("No (Số báo giá)", value=info["no"], key="mw_no")
        info["revise"] = st.text_input("Revise", value=info["revise"], key="mw_revise")
    with c2:
        info["date"] = st.date_input("Date (Ngày báo giá)", value=info["date"], key="mw_date")
        info["validity"] = st.text_input("Validity", value=info["validity"], key="mw_validity")

    with st.expander("🤝 Điều khoản thương mại"):
        info["note"] = st.text_area("Lưu ý", value=info["note"], key="mw_note")
        info["deliv1"] = st.text_input("Giao hàng (dòng 1)", value=info["deliv1"], key="mw_deliv1")
        info["deliv2"] = st.text_input("Giao hàng (dòng 2)", value=info["deliv2"], key="mw_deliv2")
        info["payment"] = st.text_area("Thanh toán", value=info["payment"], key="mw_payment")
        info["bank"] = st.text_area("Thông tin tài khoản (nếu cần)", value=info["bank"], key="mw_bank")
    st.button("➡️ Tiếp: Xuất báo giá", type="primary", on_click=_goto, args=(3,), use_container_width=True)


# ------------------------------------------------------------------ bước 4: xuất
def _step_export(c):
    df = st.session_state.get("df_editable")
    if df is None:
        st.info("Chưa có dữ liệu để xuất. Quay lại bước 1.")
        st.button("⬅️ Về bước 1", on_click=_goto, args=(0,), use_container_width=True)
        return
    df = c.recalc(df)
    info = st.session_state.m_info
    hdr = {k: info[k] for k in ["to", "add", "att", "mobile", "email", "ref", "no", "revise", "validity", "project"]}
    hdr["date"] = info["date"].strftime("%d/%m/%Y")
    ftr = {k: info[k] for k in ["note", "deliv1", "deliv2", "payment", "bank"]}

    audit_rows = [{"ma": code_of(r["Mã hàng"]), "sl": r["SL"], "don_gia": r["Đơn giá (VND)"]}
                  for _, r in df.iterrows()]
    issues = audit_quote(audit_rows, c.store)
    n_err = sum(1 for i in issues if i["level"] == "error")
    n_warn = sum(1 for i in issues if i["level"] == "warn")
    n_nocode = int(sum(1 for r in audit_rows if not r["ma"]))
    n_noprice = int(sum(1 for r in audit_rows if not r["don_gia"]))

    cong, thue, tong = c.quote_totals(df)
    st.markdown("### 📤 Xuất báo giá")
    st.markdown(f"Cộng: **{_money(cong)} đ**  \nThuế GTGT (8%): **{_money(thue)} đ**  \n"
                f"Tổng cộng: **{_money(tong)} đ**")
    if n_nocode or n_noprice:
        st.info(f"Có {n_nocode} dòng chưa ghép mã và {n_noprice} dòng chưa có giá. "
                "Các dòng này được tô vàng để bạn soát lại.")

    with st.expander(f"🩺 Kiểm tra ({n_err} lỗi · {n_warn} cảnh báo)", expanded=n_err > 0):
        if not issues:
            st.success("Không phát hiện vấn đề.")
        show = {"error": st.error, "warn": st.warning, "info": st.info}
        for w in issues[:30]:
            show[w["level"]](w["message"])
        if len(issues) > 30:
            st.caption(f"... và {len(issues) - 30} mục khác.")

    with st.expander("👁️ Xem nhanh báo giá"):
        view = pd.DataFrame({
            "STT": range(1, len(df) + 1),
            "Mã": [code_of(v) for v in df["Mã hàng"]],
            "Diễn giải": [clean(r["Diễn giải"]) if code_of(r["Mã hàng"]) else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
                          for _, r in df.iterrows()],
            "SL": df["SL"], "Đơn giá": df["Đơn giá (VND)"], "Thành tiền": df["SL"] * df["Đơn giá (VND)"],
        })
        st.dataframe(view, use_container_width=True, hide_index=True, height=300)

    cust = st.checkbox("📨 Bản gửi khách (không tô vàng)", value=False, key="m_cust")
    base = st.text_input("Tên file (không kèm đuôi)", value=f"BaoGia_TSC_{datetime.now().strftime('%Y%m%d')}",
                         key="m_out_base")
    sig = int(pd.util.hash_pandas_object(df, index=False).sum()) + hash(
        tuple(sorted((k, str(v)) for k, v in info.items())) + (cust,))

    st.caption("① Xuất Excel  →  ② tải Excel, hoặc chuyển chính file Excel đó sang PDF rồi tải.")
    if st.button("📊 ① XUẤT EXCEL (.xlsx)", use_container_width=True, key="m_btn_xl"):
        try:
            with st.spinner("Đang ghi vào mẫu Excel..."):
                st.session_state["out_xlsx"] = c.build_excel(df, hdr, ftr, customer=cust)
                st.session_state["out_xlsx_sig"] = sig
                st.session_state["out_pdf"] = None       # Excel mới thì PDF cũ không còn đúng
                st.session_state["out_pdf_sig"] = None
            c.store.record_quote_prices(audit_rows)
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi xuất Excel: {e}")
    excel_ready = bool(st.session_state.get("out_xlsx")) and st.session_state.get("out_xlsx_sig") == sig
    if st.session_state.get("out_xlsx") and not excel_ready:
        st.caption("⚠️ Dữ liệu đã thay đổi sau khi tạo file. Bấm lại 'XUẤT EXCEL' để cập nhật.")
    elif excel_ready:
        st.download_button("⬇️ TẢI FILE EXCEL", data=st.session_state["out_xlsx"], file_name=f"{base}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, key="m_dl_xl")

    if not excel_ready:
        st.button("📄 ② CHUYỂN SANG PDF", use_container_width=True, key="m_btn_pdf", disabled=True)
        st.caption("Xuất Excel trước, sau đó mới chuyển sang PDF được.")
        return
    if st.button("📄 ② CHUYỂN EXCEL SANG PDF", use_container_width=True, key="m_btn_pdf"):
        try:
            with st.spinner("Đang chuyển file Excel sang PDF..."):
                pdf_bytes = c.xlsx_to_pdf(st.session_state["out_xlsx"])
                if pdf_bytes is None:
                    st.warning("Chưa tìm thấy LibreOffice nên PDF này là bản đơn giản, chưa đúng mẫu.")
                    pdf_bytes = c.build_quote_pdf(df, hdr, ftr, customer=cust)
                st.session_state["out_pdf"] = pdf_bytes
                st.session_state["out_pdf_sig"] = st.session_state["out_xlsx_sig"]
        except ImportError:
            st.error("Chưa cài thư viện PDF. Chạy: py -m pip install reportlab")
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi chuyển PDF: {e}")
    if st.session_state.get("out_pdf") and st.session_state.get("out_pdf_sig") == sig:
        st.download_button("⬇️ TẢI FILE PDF", data=st.session_state["out_pdf"], file_name=f"{base}.pdf",
                           mime="application/pdf", use_container_width=True, key="m_dl_pdf")


# ------------------------------------------------------------------- vào đây
def run(c):
    """c: gói các hàm/dữ liệu dùng chung do app.py truyền sang (xem app.py)."""
    st.markdown(MOBILE_CSS, unsafe_allow_html=True)
    _init_state()

    if st.button("🔄 Đổi sang giao diện máy tính", key="m_switch", use_container_width=True):
        st.session_state.device = None
        st.rerun()

    # Bảng giá (dùng cho cả bước 1 và 2)
    price_items, all_sheets = [], []
    if os.path.exists(c.PRICE_LIST_PATH):
        try:
            mtime = os.path.getmtime(c.PRICE_LIST_PATH)
            names = c.cached_price_sheet_names(c.PRICE_LIST_PATH, mtime)
            all_sheets = [s for s in names if s.upper() != c.PLACEHOLDER_SHEET]
            chosen = [s for s in st.session_state.get("m_price_sel", all_sheets) if s in all_sheets] or all_sheets
            price_items = c.cached_price_items(c.PRICE_LIST_PATH, mtime, tuple(chosen))
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi đọc Bảng Giá: {e}")

    # Chuyển bước theo yêu cầu của nút bấm (phải đặt trước khi tạo ô chọn bước)
    goto = st.session_state.get("m_goto")
    if goto is not None:
        st.session_state["m_step"] = STEPS[goto]
        st.session_state["m_goto"] = None
    st.radio("Bước", STEPS, horizontal=True, key="m_step", label_visibility="collapsed")

    step = STEPS.index(st.session_state["m_step"])
    if step == 0:
        _step_boq(c, price_items, all_sheets)
    elif step == 1:
        _step_items(c, price_items)
    elif step == 2:
        _step_info()
    else:
        _step_export(c)
