"""
group_export.py - Chia báo giá thành nhiều nhóm (Van / Đầu phun / Báo cháy / Khác) và xuất MỖI NHÓM MỘT FILE
Excel (và PDF), dùng chung cho giao diện máy tính và điện thoại.

Nhóm của một dòng: ưu tiên nhóm người dùng chọn tay (cột "Nhóm"); nếu để trống thì tự suy ra từ
sheet bảng giá của mã đã ghép, chưa có mã thì suy từ chữ trong mô tả BOQ.
"""
import io
import re
import zipfile

import pandas as pd
import streamlit as st

from ai_assistant import clean, norm

GROUPS = ["Van", "Đầu phun", "Báo cháy", "Khác"]
SLUG = {"Van": "Van", "Đầu phun": "DauPhun", "Báo cháy": "BaoChay", "Khác": "Khac"}
SUFFIX = {"Van": "V", "Đầu phun": "DP", "Báo cháy": "BC", "Khác": "K"}


def _by_text(text) -> str:
    t = norm(text)
    if re.search(r"đầu phun|sprinkler|spinkler", t):
        return "Đầu phun"
    if re.search(r"báo cháy|đầu dò|chuông|còi|đèn báo|nút nhấn|trung tâm|tủ báo|bo mở rộng|ắc quy|"
                 r"\bloop\b|địa chỉ|detector|smoke|cáp tín hiệu|hochiki|module|đế đầu dò", t):
        return "Báo cháy"
    if re.search(r"\bvan\b|công tắc|đồng hồ|khớp nối|rọ hút|y lọc|bơm|ống|vòi|họng|mặt bích|"
                 r"kính quan sát|valve|switch", t):
        return "Van"
    return "Khác"


def _by_sheet(sheet) -> str:
    s = clean(sheet).upper()
    if s.startswith("VAN"):
        return "Van"
    if s.startswith(("SPINKLER", "SPRINKLER")):
        return "Đầu phun"
    if s.startswith(("HOCHIKI", "HOKICHI", "CAP", "CÁP", "THIẾT BỊ PHỤ")):
        return "Báo cháy"
    return ""


def effective_group(row, by_key) -> str:
    manual = clean(row.get("Nhóm"))
    if manual in GROUPS:
        return manual
    item = by_key.get(clean(row.get("Mã hàng")))
    if item:
        g = _by_sheet(item.get("sheet")) or _by_text(item.get("ten"))
        return g
    return _by_text(f"{clean(row.get('BOQ gốc'))} {clean(row.get('Diễn giải'))}")


def split_groups(df, by_key) -> dict:
    out = {g: [] for g in GROUPS}
    for idx, r in df.iterrows():
        out[effective_group(r, by_key)].append(idx)
    return {g: df.loc[ix].reset_index(drop=True) for g, ix in out.items() if ix}


def _zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return buf.getvalue()


def render(key, df, price_items, hdr, ftr, cust, base, sig, recalc, build_excel, xlsx_to_pdf, build_quote_pdf,
           record_prices=None):
    """Khối giao diện 'Xuất tách theo nhóm'. df: bảng đã recalc."""
    by_key = {f"{p['ma']} | {p['ten']}": p for p in price_items if p["ma"]}
    groups = split_groups(df, by_key)
    XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    rows = []
    for g, d in groups.items():
        total = float((d["SL"] * d["Đơn giá (VND)"]).sum())
        miss = int(sum(1 for _, r in d.iterrows() if not clean(r["Mã hàng"]) or float(r["Đơn giá (VND)"] or 0) <= 0))
        rows.append({"Nhóm": g, "Số dòng": len(d), "Cần soát": miss, "Cộng (VND)": f"{total:,.0f}"})
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    st.caption("Muốn chuyển dòng sang nhóm khác: đổi cột/ô **Nhóm** của dòng đó ở phần chi tiết phía trên.")

    chosen = st.multiselect("Nhóm cần xuất", list(groups), default=list(groups), key=f"{key}_sel")
    suffix = st.checkbox("Thêm hậu tố nhóm vào Số báo giá (VD: BG-141126-01-V)", value=True, key=f"{key}_sfx")
    gsig = (sig, tuple(chosen), suffix)

    if st.button("📦 XUẤT EXCEL THEO NHÓM", use_container_width=True, key=f"{key}_xl", disabled=not chosen):
        files = {}
        try:
            with st.spinner("Đang tạo file Excel cho từng nhóm..."):
                for g in chosen:
                    h = dict(hdr)
                    if suffix:
                        h["no"] = f"{hdr['no']}-{SUFFIX[g]}"
                    files[g] = build_excel(recalc(groups[g]), h, ftr, customer=cust)
            st.session_state[f"{key}_xlsx"] = files
            st.session_state[f"{key}_sig"] = gsig
            st.session_state[f"{key}_pdf"] = None
            if record_prices:
                record_prices()
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi xuất Excel theo nhóm: {e}")

    files = st.session_state.get(f"{key}_xlsx")
    if not files:
        return
    if st.session_state.get(f"{key}_sig") != gsig:
        st.caption("⚠️ Dữ liệu hoặc lựa chọn đã đổi sau khi tạo file. Bấm lại 'XUẤT EXCEL THEO NHÓM' để cập nhật.")
        return

    for g, data in files.items():
        st.download_button(f"⬇️ Excel: {g}", data=data, file_name=f"{base}_{SLUG[g]}.xlsx", mime=XL,
                           use_container_width=True, key=f"{key}_dl_{g}")
    st.download_button("🗜️ Tải tất cả Excel (.zip)",
                       data=_zip({f"{base}_{SLUG[g]}.xlsx": d for g, d in files.items()}),
                       file_name=f"{base}_Excel.zip", mime="application/zip",
                       use_container_width=True, key=f"{key}_dl_zipx")

    if st.button("📄 CHUYỂN TẤT CẢ SANG PDF", use_container_width=True, key=f"{key}_pdfbtn"):
        pdfs, simple = {}, False
        try:
            with st.spinner("Đang chuyển từng file sang PDF..."):
                for g, data in files.items():
                    p = xlsx_to_pdf(data)
                    if p is None:
                        simple = True
                        h = dict(hdr)
                        if suffix:
                            h["no"] = f"{hdr['no']}-{SUFFIX[g]}"
                        p = build_quote_pdf(recalc(groups[g]), h, ftr, customer=cust)
                    pdfs[g] = p
            st.session_state[f"{key}_pdf"] = pdfs
            st.session_state[f"{key}_pdfsimple"] = simple
        except ImportError as e:
            st.error(f"Thiếu thư viện PDF ({e}).")
        except Exception as e:  # noqa: BLE001
            st.error(f"Lỗi chuyển PDF: {e}")
    pdfs = st.session_state.get(f"{key}_pdf")
    if pdfs:
        if st.session_state.get(f"{key}_pdfsimple"):
            st.warning("Chưa tìm thấy LibreOffice nên PDF là bản đơn giản, chưa đúng mẫu.")
        for g, data in pdfs.items():
            st.download_button(f"⬇️ PDF: {g}", data=data, file_name=f"{base}_{SLUG[g]}.pdf",
                               mime="application/pdf", use_container_width=True, key=f"{key}_dlp_{g}")
        st.download_button("🗜️ Tải tất cả PDF (.zip)",
                           data=_zip({f"{base}_{SLUG[g]}.pdf": d for g, d in pdfs.items()}),
                           file_name=f"{base}_PDF.zip", mime="application/zip",
                           use_container_width=True, key=f"{key}_dl_zipp")
