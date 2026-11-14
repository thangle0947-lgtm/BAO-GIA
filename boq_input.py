"""
boq_input.py - Nhập BOQ bằng ẢNH (Claude đọc) hoặc DÁN CHỮ, dùng chung cho giao diện máy tính và điện thoại.

render(key, price_items, store, use_llm) -> (rows, stats) khi người dùng bấm "Xử lý", ngược lại None.
Kết quả đi qua đúng build_quote_rows như file Excel, nên phần ghép mã / giá / học hỏi giống hệt.
"""
import pandas as pd
import streamlit as st

from ai_assistant import (
    ITEM_COLS, build_quote_rows, extract_boq_from_image, has_llm, parse_boq_text,
)

COLS_MAP = {"name": "Tên hàng", "qty": "SL", "unit": "ĐVT", "spec": "Thông số", "price": None}


def _items_df(rows):
    return pd.DataFrame(rows, columns=ITEM_COLS)


def render(key, price_items, store, use_llm=False):
    skey = f"{key}_items"
    st.markdown("##### 📷 Nhập BOQ từ ảnh hoặc chữ")
    tab_img, tab_txt = st.tabs(["📷 Ảnh BOQ", "✍️ Dán chữ"])

    with tab_img:
        files = st.file_uploader("Chọn ảnh BOQ (chụp thẳng, rõ chữ)", type=["png", "jpg", "jpeg", "webp"],
                                 accept_multiple_files=True, key=f"{key}_img")
        if files:
            st.image([f.getvalue() for f in files], width=140)
        if not has_llm():
            st.info("Chưa bật được Claude đọc ảnh (thiếu ANTHROPIC_API_KEY hoặc thư viện anthropic). "
                    "Cách khác: trên điện thoại, mở ảnh → giữ vào chữ trong ảnh (Văn bản trực tiếp / "
                    "Live Text) → Sao chép → dán sang tab **Dán chữ**.")
        if st.button("🔍 Đọc ảnh bằng AI", key=f"{key}_read", use_container_width=True,
                     disabled=not files or not has_llm()):
            all_rows, errors = [], []
            with st.spinner("Claude đang đọc ảnh..."):
                for f in files:
                    rows, err = extract_boq_from_image(f.getvalue())
                    all_rows += rows
                    if err:
                        errors.append(f"{f.name}: {err}")
            for e in errors:
                st.error(e)
            if all_rows:
                st.session_state[skey] = _items_df(all_rows)
                st.session_state[f"{key}_ver"] = st.session_state.get(f"{key}_ver", 0) + 1

    with tab_txt:
        st.caption("Mỗi dòng một hàng hóa. Số lượng ở cuối dòng (VD: `Van cổng DN100 5 cái`), "
                   "hoặc ngăn cột bằng Tab hoặc dấu `|`. Không có số lượng thì mặc định 1, bạn sửa ở bảng bên dưới.")
        text = st.text_area("Dán chữ BOQ vào đây", height=180, key=f"{key}_text")
        if st.button("➡️ Chuyển thành bảng", key=f"{key}_parse", use_container_width=True, disabled=not text.strip()):
            st.session_state[skey] = _items_df(parse_boq_text(text))
            st.session_state[f"{key}_ver"] = st.session_state.get(f"{key}_ver", 0) + 1

    items = st.session_state.get(skey)
    if items is None or items.empty:
        return None

    st.markdown("**✏️ Kiểm tra và sửa bảng trước khi xử lý** (sửa chữ đọc sai, thêm/xóa dòng nếu cần)")
    edited = st.data_editor(
        items, key=f"{key}_editor_{st.session_state.get(f'{key}_ver', 0)}", num_rows="dynamic",
        use_container_width=True, hide_index=True,
        column_config={"Tên hàng": st.column_config.TextColumn("Tên hàng", width="large"),
                       "ĐVT": st.column_config.TextColumn("ĐVT", width="small"),
                       "SL": st.column_config.NumberColumn("SL", min_value=0, format="%g", width="small"),
                       "Thông số": st.column_config.TextColumn("Thông số")})
    st.caption(f"{len(edited)} dòng")
    if st.button("🚀 XỬ LÝ THÀNH BÁO GIÁ", type="primary", use_container_width=True, key=f"{key}_go"):
        df_raw = edited.copy()
        df_raw["Tên hàng"] = df_raw["Tên hàng"].fillna("").astype(str)
        df_raw["ĐVT"] = df_raw["ĐVT"].fillna("").astype(str)
        df_raw["SL"] = pd.to_numeric(df_raw["SL"], errors="coerce").fillna(0)
        df_raw["Thông số"] = df_raw["Thông số"].fillna("").astype(str)
        if not price_items:
            st.warning("Chưa có bảng giá nào được chọn. Các dòng sẽ không được ghép mã.")
        with st.spinner("AI đang phân tích và gợi ý mã hàng..."):
            return build_quote_rows(df_raw, dict(COLS_MAP, spec=None), price_items, store, use_llm=use_llm)
    return None
