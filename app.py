import os
import copy
import re
import unicodedata
from datetime import datetime
from io import BytesIO

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.worksheet.properties import PageSetupProperties
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
TEMPLATE_PATH = os.path.join(BASE_DIR, "MauBaoGia_1.xlsx")
if not os.path.exists(TEMPLATE_PATH):
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
for _key, _default in {"df_editable": None, "learned_keys": set(), "last_stats": None, "editor_ver": 0, "prev_ver": 0, "merge_msg": 0, "device": None}.items():
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


def merge_duplicate_rows(df):
    """Gộp các dòng đã có mã mà cùng Mã hàng + Diễn giải + Xuất xứ + ĐVT + Đơn giá thành 1 dòng, cộng dồn SL.
    Dòng chưa có mã không bị gộp. Trả về (df mới, số dòng đã gộp vào dòng khác, bảng đổi chỉ số cũ -> mới)."""
    df = df.copy()
    first_of, keep, merged = {}, [], 0
    owner = {}  # chỉ số cũ -> chỉ số cũ của dòng giữ lại
    for idx, r in df.iterrows():
        code = clean(r["Mã hàng"])
        if not code:
            keep.append(idx)
            owner[idx] = idx
            continue
        key = (code, clean(r["Diễn giải"]), clean(r["Xuất xứ"]), clean(r["ĐVT"]), float(r["Đơn giá (VND)"] or 0))
        if key in first_of:
            f = first_of[key]
            df.at[f, "SL"] = float(df.at[f, "SL"] or 0) + float(r["SL"] or 0)
            note = clean(r["Ghi chú"])
            old = clean(df.at[f, "Ghi chú"])
            if note and note not in old.split(" | "):
                df.at[f, "Ghi chú"] = f"{old} | {note}" if old else note
            owner[idx] = f
            merged += 1
        else:
            first_of[key] = idx
            keep.append(idx)
            owner[idx] = idx
    pos = {old: i for i, old in enumerate(keep)}
    mapping = {old: pos[o] for old, o in owner.items()}
    return df.loc[keep].reset_index(drop=True), merged, mapping


def merge_and_remap(df):
    """Gộp dòng trùng, đổi lại chỉ số trong tập 'đã xử lý' của phiên, và ghi lời nhắn để báo người dùng."""
    df, merged, mapping = merge_duplicate_rows(df)
    if merged:
        seen = st.session_state.learned_keys
        remapped = set()
        for k in seen:
            if isinstance(k, tuple) and k and k[0] == "ma":
                remapped.add(("ma", mapping.get(k[1], k[1]), k[2]))
            elif isinstance(k, tuple) and k and isinstance(k[0], int):
                remapped.add((mapping.get(k[0], k[0]),) + tuple(k[1:]))
            else:
                remapped.add(k)
        st.session_state.learned_keys = remapped
        st.session_state.merge_msg = st.session_state.get("merge_msg", 0) + merged
    return df, merged


def on_editor_change(editor_key, price_items, store, source="main"):
    """Chạy ngay khi người dùng sửa bảng: xóa dòng, đồng bộ mã, duyệt. Không cần bấm nút.
    source='main': bảng chi tiết phía trên; source='preview': bảng sửa ngay trong phần xem trước.
    Hai bảng dùng chung một df nên bảng còn lại phải được làm mới (đổi key) để không đè lại giá trị cũ."""
    state = st.session_state.get(editor_key)
    if not isinstance(state, dict) or st.session_state.df_editable is None:
        return
    df = recalc(st.session_state.df_editable)
    edited_rows = state.get("edited_rows") or {}
    df = apply_cell_edits(df, edited_rows)
    synced = sync_codes(df, edited_rows, price_items, store, st.session_state.learned_keys)
    applied = apply_approved(df, price_items, store, st.session_state.learned_keys)
    df, structural = remove_and_add_rows(df, state)
    merged = 0
    if synced or applied:                       # vừa duyệt/chọn mã -> gộp các dòng trùng mã + mô tả
        df, merged = merge_and_remap(recalc(df))
    st.session_state.df_editable = recalc(df)
    if source == "main":
        st.session_state.prev_ver += 1          # làm mới bảng xem trước
        if structural or synced or applied or merged:
            st.session_state.editor_ver += 1
    else:
        st.session_state.editor_ver += 1        # làm mới bảng chi tiết phía trên
        if synced or merged:
            st.session_state.prev_ver += 1      # đổi mã: nạp lại tên/giá/xuất xứ vừa điền


def ascii_key(v):
    """Chuỗi không dấu, chữ thường, để so khớp nhãn dù gõ sai dấu (ví dụ 'Thuể' / 'Thanh tóan')."""
    s = unicodedata.normalize("NFD", clean(v)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", s).strip()


VAT_RATE = 0.08


def quote_totals(df):
    """Tính Cộng, Thuế, Tổng cho bản xem trước và PDF (Excel vẫn dùng công thức của mẫu)."""
    amount = 0.0
    for _, r in df.iterrows():
        p = float(r["Đơn giá (VND)"] or 0)
        if p > 0:
            amount += float(r["SL"] or 0) * p
    vat = amount * VAT_RATE
    return amount, vat, amount + vat


def _money(v):
    return f"{v:,.0f}" if v else ""


def render_preview_html(df, h, f, split=False):
    """Bản xem trước HTML có cùng bố cục với file xuất ra."""
    import html as _h
    esc = lambda v: _h.escape(str(v or ""))
    cong, thue, tong = quote_totals(df)
    rows = []
    cell_style = "border:1px solid #6b5a22;padding:5px;color:#ECE7D8;"
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        code = code_of(r["Mã hàng"])
        price = float(r["Đơn giá (VND)"] or 0)
        name = clean(r["Diễn giải"]) if code else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
        flag = (not code) or price <= 0
        bg = "background:#FFF2A8;color:#111;" if flag else ""
        amount = float(r["SL"] or 0) * price
        cells = [(str(i), "center"), (code, "center"), (name, "left"), (clean(r["Xuất xứ"]), "center"),
                 (clean(r["ĐVT"]) or "Cái", "center"), (f"{float(r['SL'] or 0):,.2f}", "center"),
                 (_money(price), "center"), (_money(amount), "center"), ("" if _is_internal_note(clean(r["Ghi chú"])) else clean(r["Ghi chú"]), "left")]
        tds = "".join(f"<td style='{cell_style}{bg}text-align:{al}'>{esc(txt)}</td>" for txt, al in cells)
        rows.append(f"<tr>{tds}</tr>")
    th = "border:1px solid #6b5a22;padding:6px;background:#2a2412;color:#C9A227;font-weight:700;"
    td = "border:1px solid #6b5a22;padding:5px;color:#ECE7D8;text-align:center;"
    head = (
        f"<table style='width:100%;border-collapse:collapse;margin-bottom:12px;font-size:14px;color:#ECE7D8'>"
        f"<tr><td style='padding:3px 8px'><b>To:</b> {esc(h['to'])}</td><td style='padding:3px 8px'><b>No:</b> {esc(h['no'])}</td></tr>"
        f"<tr><td style='padding:3px 8px'><b>Add:</b> {esc(h['add'])}</td><td style='padding:3px 8px'><b>Date:</b> {esc(h['date'])}</td></tr>"
        f"<tr><td style='padding:3px 8px'><b>Att:</b> {esc(h['att'])}</td><td style='padding:3px 8px'><b>Validity:</b> {esc(h['validity'])}</td></tr>"
        f"<tr><td style='padding:3px 8px'><b>Mobile:</b> {esc(h['mobile'])}</td><td style='padding:3px 8px'><b>Project:</b> {esc(h['project'])}</td></tr>"
        f"<tr><td style='padding:3px 8px'><b>Email:</b> {esc(h['email'])}</td><td></td></tr>"
        f"<tr><td style='padding:3px 8px'><b>Ref:</b> {esc(h['ref'])}</td><td></td></tr></table>")
    table = (
        "<div style='overflow-x:auto'><table style='width:100%;border-collapse:collapse;font-size:13px'>"
        f"<tr><th style='{th}'>STT</th><th style='{th}'>Mã hàng</th><th style='{th}'>Diễn giải</th>"
        f"<th style='{th}'>Xuất xứ</th><th style='{th}'>ĐVT</th><th style='{th}'>SL</th>"
        f"<th style='{th}'>Đơn giá</th><th style='{th}'>Thành tiền</th><th style='{th}'>Ghi chú</th></tr>"
        + "".join(rows) + "</table></div>")
    totals = (
        "<table style='width:100%;margin-top:8px;font-size:14px;color:#ECE7D8'>"
        f"<tr><td style='text-align:right;padding:4px'><b>Cộng:</b></td><td style='width:200px;text-align:right'>{_money(cong)}</td></tr>"
        f"<tr><td style='text-align:right;padding:4px'><b>Thuế GTGT (8%):</b></td><td style='text-align:right'>{_money(thue)}</td></tr>"
        f"<tr><td style='text-align:right;padding:4px;color:#C9A227'><b>Tổng cộng (VND):</b></td>"
        f"<td style='text-align:right;color:#C9A227;font-weight:700'>{_money(tong)}</td></tr></table>")
    foot = (
        "<div style='margin-top:12px;font-size:14px;color:#ECE7D8'>"
        f"<p><b>Lưu ý:</b> {esc(f['note'])}</p>"
        f"<p><b>Giao hàng:</b> {esc(f['deliv1'])}<br>{esc(f['deliv2'])}</p>"
        f"<p><b>Thanh toán:</b><br>{esc(f['payment']).replace(chr(10), '<br>')}</p></div>")
    if split:
        return head, totals + foot
    return head + table + totals + foot


def clean_xlsx_for_excel(data: bytes, cached: dict = None) -> bytes:
    """Gỡ các phần tàn dư của mẫu (lọc tự động, chế độ xem tùy chỉnh, vùng in cũ) để Excel không báo lỗi sửa file."""
    import re as _re
    import zipfile as _zf
    src = _zf.ZipFile(BytesIO(data))
    out = BytesIO()
    with _zf.ZipFile(out, "w", _zf.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == "xl/workbook.xml":
                text = content.decode("utf-8")
                text = _re.sub(r"<definedName[^>]*wvu[^>]*>.*?</definedName>", "", text, flags=_re.S)
                text = _re.sub(r"<definedName[^>]*_FilterDatabase[^>]*>.*?</definedName>", "", text, flags=_re.S)
                text = _re.sub(r'\sfirstSheet="\d+"', "", text)  # file xuất chỉ còn 1 sheet
                content = text.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = content.decode("utf-8")
                text = _re.sub(r"<customSheetViews>.*?</customSheetViews>", "", text, flags=_re.S)
                text = _re.sub(r"<autoFilter\b[^>]*/>", "", text)
                text = _re.sub(r"<autoFilter\b[^>]*>.*?</autoFilter>", "", text, flags=_re.S)
                # Ghi sẵn kết quả công thức: điện thoại / Zalo / iOS Quick Look không tự tính nên sẽ trống nếu thiếu
                for coord, val in (cached or {}).items():
                    text = _re.sub(r'(<c r="%s"[^>]*>\s*<f>[^<]*</f>)\s*(?:<v\s*/>|<v></v>)' % coord,
                                   lambda m, v=val: f"{m.group(1)}<v>{v!r}</v>", text)
                content = text.encode("utf-8")
            dst.writestr(item, content)
    return out.getvalue()

def _fix_logo_anchor(xml: bytes) -> bytes:
    """Đổi neo 2 ô (logo giãn khi cột bị nới rộng) thành neo 1 ô với kích thước cố định."""
    import re as _re
    text = xml.decode("utf-8")

    def conv(m):
        block = m.group(0)
        frm = _re.search(r"<xdr:from>.*?</xdr:from>", block, _re.S)
        ext = _re.search(r'<a:ext cx="(\d+)" cy="(\d+)"\s*/>', block)
        if not frm or not ext:
            return block
        body = _re.sub(r"^<xdr:twoCellAnchor[^>]*>", "", block)
        body = _re.sub(r"</xdr:twoCellAnchor>$", "", body)
        body = _re.sub(r"<xdr:from>.*?</xdr:from>\s*<xdr:to>.*?</xdr:to>", "", body, flags=_re.S)
        return (f"<xdr:oneCellAnchor>{frm.group(0)}<xdr:ext cx=\"{ext.group(1)}\" cy=\"{ext.group(2)}\"/>"
                f"{body}</xdr:oneCellAnchor>")

    text = _re.sub(r"<xdr:twoCellAnchor\b.*?</xdr:twoCellAnchor>", conv, text, flags=_re.S)
    return text.encode("utf-8")


def restore_template_drawing(data: bytes, template_path: str, sheet_name: str) -> bytes:
    """openpyxl ghi lại ảnh/logo không chuẩn nên Excel báo 'Phần Đã loại bỏ: /xl/drawings/drawing1.xml'.
    Hàm này bỏ phần ảnh do openpyxl ghi và ghép NGUYÊN BẢN drawing + ảnh gốc của mẫu (do Excel tạo) vào file xuất.
    File xuất chỉ có 1 sheet (sheet1.xml)."""
    import posixpath as _pp
    import re as _re
    import zipfile as _zf
    from xml.etree import ElementTree as _ET

    REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    RID = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    try:
        tz = _zf.ZipFile(template_path)
        wbx = _ET.fromstring(tz.read("xl/workbook.xml"))
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rid = None
        for sh in wbx.find("m:sheets", ns):
            if sh.get("name") == sheet_name:
                rid = sh.get(f"{{{RID}}}id")
        if rid is None:
            return data
        wb_rels = _ET.fromstring(tz.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in wb_rels if r.get("Id") == rid)
        sheet_path = target.lstrip("/") if target.startswith("/") else _pp.normpath(_pp.join("xl", target))
        rels_path = _pp.join(_pp.dirname(sheet_path), "_rels", _pp.basename(sheet_path) + ".rels")
        if rels_path not in tz.namelist():
            return data
        drw_target = next((r.get("Target") for r in _ET.fromstring(tz.read(rels_path))
                           if r.get("Type", "").endswith("/drawing")), None)
        if not drw_target:
            return data
        drw_path = _pp.normpath(_pp.join(_pp.dirname(sheet_path), drw_target)) if not drw_target.startswith("/") \
            else drw_target.lstrip("/")
        drw_xml = tz.read(drw_path)
        drw_xml = _fix_logo_anchor(drw_xml)
        drw_rels_path = _pp.join(_pp.dirname(drw_path), "_rels", _pp.basename(drw_path) + ".rels")
        media = {}      # tên mới trong file xuất -> nội dung
        rel_items = []  # (Id, Type, Target mới)
        if drw_rels_path in tz.namelist():
            for r in _ET.fromstring(tz.read(drw_rels_path)):
                if r.get("TargetMode") == "External":
                    rel_items.append((r.get("Id"), r.get("Type"), r.get("Target"), True))
                    continue
                src = _pp.normpath(_pp.join(_pp.dirname(drw_path), r.get("Target")))
                new_name = "xl/media/" + _pp.basename(src)
                media[new_name] = tz.read(src)
                rel_items.append((r.get("Id"), r.get("Type"), "/" + new_name, False))
    except Exception:
        return data  # mẫu không có ảnh hoặc không đọc được: giữ nguyên file

    src_zip = _zf.ZipFile(BytesIO(data))
    out = BytesIO()
    with _zf.ZipFile(out, "w", _zf.ZIP_DEFLATED) as dst:
        for item in src_zip.infolist():
            name = item.filename
            if name.startswith("xl/drawings/") or name.startswith("xl/media/"):
                continue  # bỏ bản ảnh do openpyxl ghi
            content = src_zip.read(name)
            if name == "xl/worksheets/_rels/sheet1.xml.rels":
                text = content.decode("utf-8")
                text = _re.sub(r"<Relationship\b[^>]*relationships/drawing\"[^>]*/>", "", text)
                content = text.encode("utf-8")
            elif name == "xl/worksheets/sheet1.xml":
                text = content.decode("utf-8")
                text = _re.sub(r"<drawing\b[^>]*/>", "", text)
                tag = f'<drawing xmlns:r="{RID}" r:id="rIdTscDrw"/>'
                for anchor in ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls",
                               "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"):
                    pos = text.find(anchor)
                    if pos != -1:
                        text = text[:pos] + tag + text[pos:]
                        break
                content = text.encode("utf-8")
            elif name == "[Content_Types].xml":
                text = content.decode("utf-8")
                text = _re.sub(r"<Override\b[^>]*PartName=\"/xl/drawings/[^>]*/>", "", text)
                add = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                exts = {_pp.splitext(n)[1].lstrip(".").lower() for n in media}
                ctypes = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif",
                          "emf": "image/x-emf", "wmf": "image/x-wmf", "bmp": "image/bmp"}
                for e in exts:
                    if e in ctypes and f'Extension="{e}"' not in text:
                        add += f'<Default Extension="{e}" ContentType="{ctypes[e]}"/>'
                text = text.replace("</Types>", add + "</Types>")
                content = text.encode("utf-8")
            dst.writestr(name, content)
        if "xl/worksheets/_rels/sheet1.xml.rels" not in src_zip.namelist():
            dst.writestr("xl/worksheets/_rels/sheet1.xml.rels",
                         f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL}"></Relationships>')
        dst.writestr("xl/drawings/drawing1.xml", drw_xml)
        rels = "".join(
            f'<Relationship Id="{i}" Type="{t}" Target="{tg}"' + (' TargetMode="External"' if ext else "") + "/>"
            for i, t, tg, ext in rel_items)
        dst.writestr("xl/drawings/_rels/drawing1.xml.rels",
                     f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{REL}">{rels}</Relationships>')
        for n, c in media.items():
            dst.writestr(n, c)
    # thêm quan hệ drawing vào sheet1.xml.rels (sau khi đã chắc chắn file rels tồn tại)
    out2 = BytesIO()
    zin = _zf.ZipFile(BytesIO(out.getvalue()))
    with _zf.ZipFile(out2, "w", _zf.ZIP_DEFLATED) as dst:
        for item in zin.infolist():
            c = zin.read(item.filename)
            if item.filename == "xl/worksheets/_rels/sheet1.xml.rels":
                t = c.decode("utf-8")
                t = t.replace("</Relationships>",
                              f'<Relationship Id="rIdTscDrw" Type="{RID}/drawing" Target="/xl/drawings/drawing1.xml"/></Relationships>')
                c = t.encode("utf-8")
            dst.writestr(item.filename, c)
    return out2.getvalue()


def find_libreoffice():
    import shutil as _sh
    for name in ("soffice", "libreoffice"):
        p = _sh.which(name)
        if p:
            return p
    for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        if os.path.exists(p):
            return p
    return None


def xlsx_to_pdf(xlsx_bytes: bytes):
    """Chuyển file Excel đã điền thành PDF đúng mẫu bằng LibreOffice. Trả về None nếu không có LibreOffice."""
    import subprocess
    import tempfile
    exe = find_libreoffice()
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "baogia.xlsx")
        with open(src, "wb") as fh:
            fh.write(xlsx_bytes)
        # Hồ sơ riêng + HOME ghi được: bắt buộc khi chạy trên server (Streamlit Cloud/Linux)
        profile = "file://" + os.path.join(tmp, "lo_profile").replace("\\", "/")
        env = dict(os.environ, HOME=tmp)
        try:
            subprocess.run([exe, f"-env:UserInstallation={profile}", "--headless", "--norestore",
                            "--convert-to", "pdf", "--outdir", tmp, src],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180,
                           check=False, env=env)
        except (subprocess.TimeoutExpired, OSError):
            return None
        out = os.path.join(tmp, "baogia.pdf")
        if not os.path.exists(out):
            return None
        with open(out, "rb") as fh:
            return fh.read()


def build_quote_pdf(df, h, f, customer: bool = False) -> bytes:
    """Xuất PDF báo giá cùng nội dung với file Excel. Cần thư viện reportlab."""
    from io import BytesIO as _BIO
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    reg, bold = "Helvetica", "Helvetica-Bold"
    for r_path, b_path in [
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]:
        if os.path.exists(r_path):
            pdfmetrics.registerFont(TTFont("QF", r_path))
            pdfmetrics.registerFont(TTFont("QFB", b_path if os.path.exists(b_path) else r_path))
            reg, bold = "QF", "QFB"
            break

    gold = colors.HexColor("#9C7A12")
    body = ParagraphStyle("body", fontName=reg, fontSize=8.5, leading=11)
    small = ParagraphStyle("small", fontName=reg, fontSize=7.5, leading=9)
    title = ParagraphStyle("title", fontName=bold, fontSize=16, leading=20, alignment=1, textColor=gold)
    cell_bold = ParagraphStyle("cb", fontName=bold, fontSize=8.5, leading=11, alignment=1)

    story = []
    if os.path.exists(LOGO_PATH):
        story.append(Image(LOGO_PATH, width=95 * mm, height=95 * mm * 331 / 900))
    story.append(Spacer(1, 4))
    story.append(Paragraph("BÁO GIÁ", title))
    story.append(Spacer(1, 6))

    left = [("To:", h["to"]), ("Add:", h["add"]), ("Att:", h["att"]), ("Mobile:", h["mobile"]),
            ("Email:", h["email"]), ("Ref:", h["ref"])]
    right = [("No:", h["no"]), ("Date:", h["date"]), ("Revise:", h["revise"]),
             ("Validity:", h["validity"]), ("Project:", h["project"])]
    hdr_rows = []
    for i in range(max(len(left), len(right))):
        l = left[i] if i < len(left) else ("", "")
        rr = right[i] if i < len(right) else ("", "")
        hdr_rows.append([Paragraph(f"<b>{l[0]}</b> {l[1] or ''}", body), Paragraph(f"<b>{rr[0]}</b> {rr[1] or ''}", body)])
    ht = Table(hdr_rows, colWidths=[90 * mm, 90 * mm])
    ht.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    story += [ht, Spacer(1, 8)]

    data = [[Paragraph(t, cell_bold) for t in ["STT", "Mã hàng", "Diễn giải", "Xuất xứ", "ĐVT", "SL",
                                                "Đơn giá\n(VND)", "Thành tiền\n(VND)", "Ghi chú"]]]
    flags = []
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        code = code_of(r["Mã hàng"])
        price = float(r["Đơn giá (VND)"] or 0)
        name = clean(r["Diễn giải"]) if code else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
        amount = float(r["SL"] or 0) * price
        if (not customer) and ((not code) or price <= 0):
            flags.append(len(data))
        data.append([Paragraph(str(i), body), Paragraph(code, body), Paragraph(name, body),
                     Paragraph(clean(r["Xuất xứ"]), body), Paragraph(clean(r["ĐVT"]) or "Cái", body),
                     Paragraph(f"{float(r['SL'] or 0):,.2f}", body),
                     Paragraph(_money(price), body), Paragraph(_money(amount), body),
                     Paragraph("" if _is_internal_note(clean(r["Ghi chú"])) else clean(r["Ghi chú"]), small)])
    widths = [9, 22, 36, 21, 10, 18, 20, 26, 18]  # tổng 180 mm = khổ A4 trừ lề
    tbl = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
    st_list = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9C7A12")),
               ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
               ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F3EBCF")),
               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
    for ridx in flags:
        st_list.append(("BACKGROUND", (0, ridx), (-1, ridx), colors.HexColor("#FFF2A8")))
    tbl.setStyle(TableStyle(st_list))
    story += [tbl, Spacer(1, 6)]

    cong, thue, tong = quote_totals(df)
    tot = Table([[Paragraph("<b>Cộng:</b>", body), Paragraph(_money(cong), body)],
                 [Paragraph("<b>Thuế GTGT (8%):</b>", body), Paragraph(_money(thue), body)],
                 [Paragraph("<b>Tổng cộng (VND):</b>", body), Paragraph(f"<b>{_money(tong)}</b>", body)]],
                colWidths=[60 * mm, 40 * mm])
    tot.setStyle(TableStyle([("ALIGN", (1, 0), (1, -1), "RIGHT"), ("LINEABOVE", (0, 2), (-1, 2), 0.8, gold)]))
    wrap = Table([[None, tot]], colWidths=[80 * mm, 100 * mm])
    story += [wrap, Spacer(1, 10)]

    story.append(Paragraph(f"<b>Lưu ý:</b> {f['note']}", body))
    story.append(Paragraph(f"<b>Giao hàng:</b> {f['deliv1']}", body))
    if f.get("deliv2"):
        story.append(Paragraph(f["deliv2"], body))
    story.append(Paragraph("<b>Thanh toán:</b>", body))
    for line in (f["payment"] or "").split("\n"):
        if line.strip():
            story.append(Paragraph(line.strip(), body))
    if f.get("bank"):
        story.append(Spacer(1, 4))
        story.append(Paragraph(f["bank"].replace("\n", "<br/>"), body))
    story += [Spacer(1, 14), Paragraph("<b>Người báo giá</b>", body)]

    buf = _BIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm, title="Báo giá")
    doc.build(story)
    return buf.getvalue()


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


def _is_internal_note(t: str) -> bool:
    """Ghi chú nội bộ do AI sinh (không được đi vào file gửi khách)."""
    k = t.lower()
    return any(p in k for p in ("đã ghép tay", "độ khớp", "claude đề xuất", "đã học", "đã duyệt", "lần trước đây"))


def build_excel(data: pd.DataFrame, hdr: dict, ftr: dict, customer: bool = False) -> bytes:
    """hdr: to, add, att, mobile, email, ref, no, date (chuỗi dd/mm/yyyy), revise, validity, project.
    ftr: note, deliv1, deliv2, payment."""
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb[QUOTE_SHEET] if QUOTE_SHEET in wb.sheetnames else wb.active
    # Chỉ giữ sheet báo giá: bỏ các sheet mẫu/cũ (ví dụ VN_1 chứa báo giá của khách khác)
    for _name in list(wb.sheetnames):
        if _name != ws.title:
            del wb[_name]
    wb.active = 0

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
    safe_write(6, 2, hdr["to"])
    safe_write(7, 2, hdr["add"])
    safe_write(8, 2, hdr["att"])
    safe_write(9, 2, hdr["mobile"])
    safe_write(10, 2, hdr["email"])
    safe_write(11, 2, hdr["ref"])
    safe_write(6, 7, hdr["no"])
    safe_write(7, 7, hdr["date"])
    safe_write(8, 7, hdr["revise"])
    safe_write(9, 7, hdr["validity"])
    safe_write(10, 7, hdr["project"])

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
    item_h = ws.row_dimensions[start_row].height or 43.9  # chiều cao dòng hàng hóa của mẫu
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

    # Mọi dòng hàng hóa cao bằng nhau (dòng thêm mới không có chiều cao sẽ bị thấp hơn dòng của mẫu)
    for r in range(start_row, start_row + n):
        ws.row_dimensions[r].height = item_h

    # 4) Ghi hàng hóa (nền vùng hàng hóa về trắng như mẫu; chỉ tô vàng dòng cần soát ở bản nội bộ)
    flag_fill = PatternFill("solid", fgColor="FFF2A8")
    cached = {}
    for r in range(start_row, start_row + n):
        ws.row_dimensions[r].fill = PatternFill(fill_type=None)
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)
    for i, (_, r) in enumerate(data.iterrows()):
        rr = start_row + i
        code = code_of(r["Mã hàng"])
        price = float(r["Đơn giá (VND)"] or 0)
        name = clean(r["Diễn giải"]) if code else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
        notes = []
        if clean(r["Ghi chú"]) and not _is_internal_note(clean(r["Ghi chú"])):
            notes.append(clean(r["Ghi chú"]))
        cached[f"H{rr}"] = float(r["SL"] or 0) * price
        safe_write(rr, 1, i + 1)
        safe_write(rr, 2, code)
        safe_write(rr, 3, name)
        safe_write(rr, 4, clean(r["Xuất xứ"]))
        safe_write(rr, 5, clean(r["ĐVT"]) or "Cái")
        safe_write(rr, 6, float(r["SL"]))
        safe_write(rr, 7, price if price > 0 else None)
        safe_write(rr, 8, f"=F{rr}*G{rr}")
        safe_write(rr, 9, " | ".join(notes))
        if not customer and (not code or price <= 0):
            for c in range(1, 10):
                ws.cell(row=rr, column=c).fill = flag_fill

    # 5) Công thức Cộng, Thuế, Tổng cộng: viết lại theo vị trí mới
    cong = start_row + n
    safe_write(cong, 8, f"=SUM(H{start_row}:H{cong - 1})")
    _sum = sum(v for v in cached.values())
    cached[f"H{cong}"] = _sum
    thue = find_row(["thue"], cong + 1, cong + 3)
    tong = find_row(["tong cong"], cong + 1, cong + 6)
    if thue is not None:
        rate = ws.cell(row=thue, column=4).value
        if isinstance(rate, (int, float)):
            safe_write(thue, 8, f"=H{cong}*D{thue}")
        else:
            safe_write(thue, 4, 0.08)
            safe_write(thue, 8, f"=H{cong}*D{thue}")
    _vat = 0.0
    if thue is not None:
        _rate = ws.cell(row=thue, column=4).value
        _vat = _sum * (_rate if isinstance(_rate, (int, float)) else 0.08)
        cached[f"H{thue}"] = _vat
    if tong is not None:
        safe_write(tong, 8, f"=H{cong}+H{thue}" if thue is not None else f"=H{cong}")
        cached[f"H{tong}"] = _sum + _vat

    # 6) Điều khoản dưới bảng
    footer_start = (tong or cong) + 1
    for r in range(footer_start, ws.max_row + 1):
        k2 = ascii_key(ws.cell(row=r, column=2).value)
        k3 = ascii_key(ws.cell(row=r, column=3).value)
        if k2.startswith("luu y"):
            safe_write(r, 3, ftr["note"])
        elif k2.startswith("giao hang"):
            safe_write(r, 3, ftr["deliv1"])
        elif k3.startswith("nhap hang") or "giay luu thong" in k3:
            safe_write(r, 3, ftr["deliv2"])
        elif k2.startswith("thanh toan"):
            # Khối thanh toán của mẫu: gom hết các dòng trống phía dưới nhãn, đến dòng "Thanh toán bằng..."
            lines = [ln.strip() for ln in ftr["payment"].split("\n") if ln.strip()]
            slots = [r]
            rr = r + 1
            while rr <= ws.max_row:
                if ascii_key(ws.cell(row=rr, column=3).value).startswith("thanh toan bang") \
                        or clean(ws.cell(row=rr, column=2).value):
                    break
                slots.append(rr)
                rr += 1
            if len(lines) >= len(slots):
                for k_, srow in enumerate(slots[:-1]):
                    safe_write(srow, 3, lines[k_])
                safe_write(slots[-1], 3, " / ".join(lines[len(slots) - 1:]) if lines else None)
            else:
                for k_, srow in enumerate(slots):
                    safe_write(srow, 3, lines[k_] if k_ < len(lines) else None)
            break

    # Mẫu có số điện thoại lặp: ô tên người báo giá đã ghi sẵn SĐT, dòng ngay dưới lại ghi thêm lần nữa -> xóa
    for r in range((tong or cong) + 1, ws.max_row):
        if "le quoc" in ascii_key(ws.cell(row=r, column=1).value):
            below = ws.cell(row=r + 1, column=1)
            if re.fullmatch(r"[\d\s.+-]{8,}", clean(below.value) or "x"):
                safe_write(r + 1, 1, None)
            break

    # Xóa nền tô ở các dòng trống phía dưới chữ ký (tránh ô đen/lạ ở cuối file)
    last_content = max((r for r in range(1, ws.max_row + 1)
                        if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, 10))),
                       default=1)
    for r in range(last_content + 1, ws.max_row + 1):
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)

    # Nới nhẹ 3 cột số để không hiện "########" trong Excel (chỉ tăng, không thu hẹp)
    for _col, _min_w in (("F", 9.5), ("G", 14.5), ("H", 16.5)):
        _cur = ws.column_dimensions[_col].width or 0
        if _cur < _min_w:
            ws.column_dimensions[_col].width = _min_w

    # Vùng in phải phủ hết phần chân báo giá sau khi đã thêm/bớt dòng
    ws.print_area = f"A1:I{last_content}"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # Không để openpyxl ghi ảnh (Excel báo lỗi drawing1.xml); ghép lại logo gốc của mẫu sau khi lưu
    ws._images = []
    buf = BytesIO()
    wb.save(buf)
    fixed = restore_template_drawing(buf.getvalue(), TEMPLATE_PATH, ws.title)
    return clean_xlsx_for_excel(fixed, cached)


# --- CHỌN THIẾT BỊ: máy tính (giao diện cũ) hoặc điện thoại (giao diện tối ưu) ---
if st.session_state.get("device") is None:
    st.markdown("""<style>
        [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"] { display: none; }
        .st-key-dev_pc button, .st-key-dev_mobile button { min-height: 96px; font-size: 20px; }
    </style>""", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center'>Bạn đang dùng thiết bị nào?</h3>", unsafe_allow_html=True)
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("🖥️  MÁY TÍNH", use_container_width=True, key="dev_pc"):
            st.session_state.device = "desktop"
            st.rerun()
    with _c2:
        if st.button("📱  ĐIỆN THOẠI", use_container_width=True, key="dev_mobile"):
            st.session_state.device = "mobile"
            st.rerun()
    st.caption("Máy tính: giao diện bảng rộng. Điện thoại: giao diện một cột, thẻ từng dòng, nút to.")
    st.stop()

if st.session_state.device == "mobile":
    import types as _types
    import mobile_ui
    mobile_ui.run(_types.SimpleNamespace(
        store=store, PRICE_LIST_PATH=PRICE_LIST_PATH, PLACEHOLDER_SHEET=PLACEHOLDER_SHEET,
        cached_price_items=cached_price_items, cached_price_sheet_names=cached_price_sheet_names,
        cached_boq_sheet_names=cached_boq_sheet_names, cached_boq_raw=cached_boq_raw,
        recalc=recalc, build_excel=build_excel, xlsx_to_pdf=xlsx_to_pdf, build_quote_pdf=build_quote_pdf,
        quote_totals=quote_totals, merge_and_remap=merge_and_remap, GOLD=GOLD))
    st.stop()

# --- SIDEBAR: TẢI DỮ LIỆU ---
if st.sidebar.button("🔄 Đổi thiết bị (máy tính / điện thoại)", use_container_width=True, key="dev_switch"):
    st.session_state.device = None
    st.rerun()
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
                    df, _ = merge_and_remap(recalc(df))
                    st.session_state.df_editable = recalc(df)
                    st.session_state.editor_ver += 1
                    st.rerun()

    # Điền mã cho các dòng đã tích Duyệt nhưng chưa được điền (kể cả khi bấm Duyệt tất cả)
    if apply_approved(df, price_items, store, st.session_state.learned_keys):
        df, _ = merge_and_remap(recalc(df))
        st.session_state.df_editable = recalc(df)
        st.session_state.editor_ver += 1
        st.rerun()

    # --- Duyệt nhanh ---
    if st.session_state.merge_msg:
        st.success(f"🔗 Đã gộp {st.session_state.merge_msg} dòng trùng mã + mô tả vào dòng đầu tiên, số lượng đã cộng dồn.")
        st.session_state.merge_msg = 0
    col_act1, col_act2, _ = st.columns([1, 1, 2])
    with col_act1:
        if st.button("✅ DUYỆT TẤT CẢ ĐỀ XUẤT", use_container_width=True):
            df.loc[df["Mã đề xuất"] != "", "Duyệt"] = True
            st.session_state.df_editable = df
            st.session_state.editor_ver += 1
            st.rerun()
    with col_act2:
        if st.button("🔗 GỘP DÒNG TRÙNG", use_container_width=True,
                     help="Các dòng đã có mã, cùng mã + mô tả + giá sẽ gộp thành 1 dòng, SL cộng dồn."):
            df2, n_merged = merge_and_remap(df)
            if n_merged:
                st.session_state.df_editable = recalc(df2)
                st.session_state.editor_ver += 1
                st.rerun()
            else:
                st.info("Không có dòng nào trùng để gộp.")
    st.caption("Tích ô 🗑️ Xóa là dòng được xóa ngay. Tích ✅ Duyệt là mã đề xuất được điền vào file. "
               "Dòng đã có mã mà trùng mã + mô tả sẽ tự gộp thành 1 và cộng SL.")

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
        f_bank = st.text_area("Thông tin tài khoản (in dưới phần thanh toán, nếu cần):", value="", key="f_bank")
        f_payment = st.text_area("Thanh toán:",
                                 value="Tạm ứng 30% khi đặt hàng\nThanh toán 70% còn lại trước khi giao hàng",
                                 key="f_payment")

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Xuất Excel (tìm dòng theo nhãn, tự thêm/bớt dòng, viết lại công thức tổng) ---

    # --- Xem trước báo giá (cập nhật ngay khi sửa ở trên) ---
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3>👁️ XEM TRƯỚC BÁO GIÁ</h3>", unsafe_allow_html=True)
    hdr_vals = {"to": h_to, "add": h_add, "att": h_att, "mobile": h_mobile, "email": h_email, "ref": h_ref,
                "no": h_no, "date": h_date.strftime("%d/%m/%Y"), "revise": h_revise,
                "validity": h_validity, "project": h_project}
    ftr_vals = {"note": f_note, "deliv1": f_deliv1, "deliv2": f_deliv2, "payment": f_payment,
                "bank": f_bank}
    st.caption("✏️ Có thể sửa trực tiếp ngay trong bảng này: chọn lại **Mã hàng** (tên, giá, xuất xứ tự điền), "
               "sửa **Đơn giá**, **SL**, **ĐVT**, **Ghi chú**. Cột Diễn giải và Thành tiền tự cập nhật.")
    head_html, tail_html = render_preview_html(df, hdr_vals, ftr_vals, split=True)
    pv = pd.DataFrame({
        "STT": range(1, len(df) + 1),
        "Mã hàng": df["Mã hàng"],
        "Diễn giải": [clean(r["Diễn giải"]) if code_of(r["Mã hàng"]) else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
                      for _, r in df.iterrows()],
        "Xuất xứ": df["Xuất xứ"], "ĐVT": df["ĐVT"], "SL": df["SL"],
        "Đơn giá (VND)": df["Đơn giá (VND)"], "Thành tiền (VND)": df["SL"] * df["Đơn giá (VND)"],
        "Ghi chú": df["Ghi chú"],
    }, index=df.index)
    pv_key = f"preview_editor_{st.session_state.editor_ver}_{st.session_state.prev_ver}"
    with st.container(border=True):
        st.markdown(head_html, unsafe_allow_html=True)
        st.data_editor(
            pv, key=pv_key, num_rows="fixed", use_container_width=True, hide_index=True,
            height=min(640, 38 + 35 * max(len(pv), 1)),
            on_change=on_editor_change, args=(pv_key, price_items, store, "preview"),
            column_order=["STT", "Mã hàng", "Diễn giải", "Xuất xứ", "ĐVT", "SL",
                          "Đơn giá (VND)", "Thành tiền (VND)", "Ghi chú"],
            column_config={
                "STT": st.column_config.NumberColumn("STT", disabled=True, width="small"),
                "Mã hàng": st.column_config.SelectboxColumn(
                    "Mã hàng", options=code_options, width="medium",
                    help="Chọn mã + kích thước từ bảng giá. Tên và giá sẽ tự điền."),
                "Diễn giải": st.column_config.TextColumn("Diễn giải", disabled=True, width="large"),
                "Xuất xứ": st.column_config.TextColumn("Xuất xứ", width="small"),
                "ĐVT": st.column_config.TextColumn("ĐVT", width="small"),
                "SL": st.column_config.NumberColumn("SL", format="%.2f", min_value=0, width="small"),
                "Đơn giá (VND)": st.column_config.NumberColumn("Đơn giá", format="%,.0f", min_value=0),
                "Thành tiền (VND)": st.column_config.NumberColumn("Thành tiền", format="%,.0f", disabled=True),
                "Ghi chú": st.column_config.TextColumn("Ghi chú"),
            })
        n_flag = int(sum(1 for _, r in df.iterrows() if (not code_of(r["Mã hàng"])) or float(r["Đơn giá (VND)"] or 0) <= 0))
        if n_flag:
            st.warning(f"⚠️ {n_flag} dòng chưa có mã hoặc chưa có giá (sẽ tô vàng trong file xuất).")
        st.markdown(tail_html, unsafe_allow_html=True)

    # --- Xuất báo giá: chọn một trong hai ---
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("<h3>📤 XUẤT BÁO GIÁ</h3>", unsafe_allow_html=True)
    n_nocode = int(sum(1 for r in audit_rows if not r["ma"]))
    n_noprice = int(sum(1 for r in audit_rows if not r["don_gia"]))
    if n_nocode or n_noprice:
        st.info(f"Có {n_nocode} dòng chưa ghép mã (ghi tên BOQ gốc) và {n_noprice} dòng chưa có giá (để trống). "
                "Các dòng này được tô vàng để bạn soát lại.")
    cust_mode = st.checkbox("📨 Bản gửi khách (không ghi chú nội bộ, không tô vàng)", value=False, key="cust_mode",
                            help="Dùng khi gửi nhanh cho khách. Bản nội bộ có ghi chú CHƯA GHÉP MÃ / CHƯA CÓ GIÁ để bạn soát.")
    base_name = st.text_input("Tên file (không kèm đuôi):",
                              value=f"BaoGia_TSC_{datetime.now().strftime('%Y%m%d')}", key="out_base")
    sig = int(pd.util.hash_pandas_object(df, index=False).sum()) + hash((hdr_vals.get("no"), ftr_vals.get("note"), cust_mode))
    st.caption("Cách làm: ① bấm XUẤT EXCEL  ②  tải Excel, hoặc bấm CHUYỂN SANG PDF để đổi chính file Excel đó thành PDF rồi tải.")
    c_xl, c_pdf = st.columns(2)
    with c_xl:
        if st.button("📊 ① XUẤT EXCEL (.xlsx)", use_container_width=True, key="btn_xl"):
            try:
                with st.spinner("Đang ghi vào mẫu Excel..."):
                    st.session_state["out_xlsx"] = build_excel(df, hdr_vals, ftr_vals, customer=cust_mode)
                    st.session_state["out_xlsx_sig"] = sig
                    st.session_state["out_pdf"] = None       # Excel mới thì PDF cũ không còn đúng
                    st.session_state["out_pdf_sig"] = None
                store.record_quote_prices(audit_rows)
            except Exception as e:
                st.error(f"Lỗi xuất Excel: {e}")
        if st.session_state.get("out_xlsx") and st.session_state.get("out_xlsx_sig") != sig:
            st.caption("⚠️ Dữ liệu đã thay đổi sau khi tạo file. Bấm lại 'XUẤT EXCEL' để cập nhật.")
        elif st.session_state.get("out_xlsx"):
            st.download_button("⬇️ TẢI FILE EXCEL", data=st.session_state["out_xlsx"],
                               file_name=f"{base_name}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, key="dl_xl")
    with c_pdf:
        excel_ready = bool(st.session_state.get("out_xlsx")) and st.session_state.get("out_xlsx_sig") == sig
        if not excel_ready:
            st.button("📄 ② CHUYỂN SANG PDF", use_container_width=True, key="btn_pdf", disabled=True,
                      help="Hãy xuất Excel trước.")
            st.caption("Xuất Excel trước, sau đó mới chuyển sang PDF được.")
        else:
            if st.button("📄 ② CHUYỂN EXCEL SANG PDF", use_container_width=True, key="btn_pdf"):
                try:
                    with st.spinner("Đang chuyển file Excel sang PDF..."):
                        pdf_bytes = xlsx_to_pdf(st.session_state["out_xlsx"])
                        if pdf_bytes is None:
                            st.warning("Chưa tìm thấy LibreOffice nên PDF này là bản đơn giản, chưa đúng mẫu. "
                                       "Cài LibreOffice (miễn phí) để chuyển PDF đúng như file Excel.")
                            pdf_bytes = build_quote_pdf(df, hdr_vals, ftr_vals, customer=cust_mode)
                        st.session_state["out_pdf"] = pdf_bytes
                        st.session_state["out_pdf_sig"] = st.session_state["out_xlsx_sig"]
                except ImportError as e:
                    st.error(f"Thiếu thư viện PDF ({e}). Kiểm tra requirements.txt có reportlab và pillow.")
                except Exception as e:
                    st.error(f"Lỗi chuyển PDF: {e}")
            if st.session_state.get("out_pdf") and st.session_state.get("out_pdf_sig") == sig:
                st.download_button("⬇️ TẢI FILE PDF", data=st.session_state["out_pdf"],
                                   file_name=f"{base_name}.pdf", mime="application/pdf",
                                   use_container_width=True, key="dl_pdf")
