"""
quote_core.py - Xử lý báo giá (không có giao diện).
Tách từ app cũ để giao diện mới dùng lại. Mọi hàm nhận dữ liệu qua tham số.
"""
import os
import re
import copy
import unicodedata
from io import BytesIO

import openpyxl
import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl.worksheet.properties import PageSetupProperties

from ai_assistant import COLUMNS, clean, code_of, item_key

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = next(
    (p for p in (os.path.join(BASE_DIR, n) for n in
                 ("MauBaoGia_34e689.xlsx", "MauBaoGia_1.xlsx", "MauBaoGia.xlsx"))
     if os.path.exists(p)),
    os.path.join(BASE_DIR, "MauBaoGia_34e689.xlsx"))
QUOTE_SHEET = "BÁO GIÁ"
LOGO_PATH = os.path.join(BASE_DIR, "tsc_logo.png")
COLUMNS_QUOTE = ["Xóa", "Duyệt", "BOQ gốc", "Mã đề xuất", "Độ khớp", "Nguồn",
                 "Mã hàng", "Diễn giải", "Xuất xứ", "ĐVT", "SL",
                 "Đơn giá (VND)", "Thành tiền (VND)", "Ghi chú", "Gợi ý"]

VAT_RATE = 0.08


def ascii_key(v):
    """Chuỗi không dấu, chữ thường, để so khớp nhãn dù gõ sai dấu (ví dụ 'Thuể' / 'Thanh tóan')."""
    s = unicodedata.normalize("NFD", clean(v)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", s).strip()


def _zip_read(z, name):
    try:
        return z.read(name)
    except KeyError:
        return None


def clean_xlsx_for_excel(data: bytes) -> bytes:
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
                n_sheets = len(_re.findall(r"<sheet\b", text))
                text = _re.sub(r'firstSheet="(\d+)"',
                               lambda m: 'firstSheet="0"' if int(m.group(1)) >= n_sheets else m.group(0), text)
                text = _re.sub(r"<definedName[^>]*wvu[^>]*>.*?</definedName>", "", text, flags=_re.S)
                text = _re.sub(r"<definedName[^>]*_FilterDatabase[^>]*>.*?</definedName>", "", text, flags=_re.S)
                content = text.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                text = content.decode("utf-8")
                text = _re.sub(r"<customSheetViews>.*?</customSheetViews>", "", text, flags=_re.S)
                text = _re.sub(r"<autoFilter\b[^>]*/>", "", text)
                text = _re.sub(r"<autoFilter\b[^>]*>.*?</autoFilter>", "", text, flags=_re.S)
                # Ô chuỗi rỗng (t="inlineStr" không có <is>) làm Excel báo "sửa file"
                text = _re.sub(r'<c ([^>]*?) t="inlineStr"\s*></c>', r"<c \1/>", text)
                text = _re.sub(r'<c ([^>]*?) t="inlineStr"\s*/>', r"<c \1/>", text)
                text = _re.sub(r"<(oddHeader|oddFooter|evenHeader|evenFooter|firstHeader|firstFooter)>\s*</\1>",
                               "", text)
                content = text.encode("utf-8")
            dst.writestr(item, content)
    return out.getvalue()


def restore_template_drawing(out_bytes: bytes) -> bytes:
    """Thay phần hình vẽ/logo do openpyxl viết lại bằng bản gốc của mẫu (Excel luôn chấp nhận bản gốc)."""
    import posixpath as _pp
    import re as _re
    import zipfile as _zf
    if not os.path.exists(TEMPLATE_PATH):
        return out_bytes
    tz = _zf.ZipFile(TEMPLATE_PATH)
    oz = _zf.ZipFile(BytesIO(out_bytes))
    if any(n.startswith("xl/media/mau_") for n in oz.namelist()):
        return out_bytes  # đã gắn logo rồi, không gắn lại (tránh trùng file trong zip)

    def sheet_part(z, sheet_name):
        wb = _zip_read(z, "xl/workbook.xml").decode("utf-8")
        m = _re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % _re.escape(sheet_name), wb)
        if not m:
            return None
        rels = _zip_read(z, "xl/_rels/workbook.xml.rels").decode("utf-8")
        t = _re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels) or \
            _re.search(r'Target="([^"]+)"[^>]*Id="%s"' % m.group(1), rels)
        return "xl/" + t.group(1).lstrip("/").replace("xl/", "") if t else None

    def drawing_of(z, sheet_part_name):
        srel = _pp.join(_pp.dirname(sheet_part_name), "_rels", _pp.basename(sheet_part_name) + ".rels")
        rels = _zip_read(z, srel)
        if not rels:
            return None
        m = _re.search(r'Target="([^"]*drawing[^"]*\.xml)"', rels.decode("utf-8"))
        if not m:
            return None
        return _pp.normpath(_pp.join(_pp.dirname(sheet_part_name), m.group(1))).lstrip("/")

    t_sheet = sheet_part(tz, QUOTE_SHEET)
    o_sheet = "xl/worksheets/sheet1.xml"
    t_draw = drawing_of(tz, t_sheet) if t_sheet else None
    o_draw = drawing_of(oz, o_sheet)
    if not t_draw or not o_draw:
        return out_bytes

    t_drawing_xml = _zip_read(tz, t_draw)
    t_drawing_rels = _zip_read(tz, _pp.join(_pp.dirname(t_draw), "_rels", _pp.basename(t_draw) + ".rels"))
    if t_drawing_xml is None:
        return out_bytes

    new_files = {}
    rels_text = t_drawing_rels.decode("utf-8") if t_drawing_rels else \
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

    def _media_sub(m):
        target = m.group(1)
        src = _pp.normpath(_pp.join(_pp.dirname(t_draw), target)).lstrip("/")
        data = _zip_read(tz, src)
        if data is None:
            return m.group(0)
        name = f"xl/media/mau_{_pp.basename(src)}"
        new_files[name] = data
        return m.group(0).replace(target, "/" + name)
    rels_text = _re.sub(r'Target="([^"]+\.(?:png|jpe?g|gif|emf|wmf))"', _media_sub, rels_text)

    out = BytesIO()
    with _zf.ZipFile(out, "w", _zf.ZIP_DEFLATED) as dst:
        for item in oz.infolist():
            if item.filename == o_draw:
                dst.writestr(item, t_drawing_xml)
            elif item.filename == _pp.join(_pp.dirname(o_draw), "_rels", _pp.basename(o_draw) + ".rels"):
                dst.writestr(item, rels_text.encode("utf-8"))
            elif item.filename.startswith("xl/media/") and item.filename not in new_files and \
                    any(item.filename.endswith(e) for e in (".png", ".jpg", ".jpeg", ".gif", ".emf", ".wmf")):
                continue  # ảnh cũ do openpyxl ghi, đã thay bằng ảnh gốc của mẫu
            else:
                dst.writestr(item, oz.read(item.filename))
        for name, data in new_files.items():
            dst.writestr(name, data)
    return out.getvalue()


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
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, src],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180, check=False)
        out = os.path.join(tmp, "baogia.pdf")
        if not os.path.exists(out):
            return None
        with open(out, "rb") as fh:
            return fh.read()


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


def render_preview_html(df, h, f):
    """Bản xem trước HTML có cùng bố cục với file xuất ra."""
    import html as _h
    esc = lambda v: _h.escape(str(v or ""))
    cong, thue, tong = quote_totals(df)
    rows = []
    cell_style = "border:1px solid #9FB3C8;padding:5px;color:#1F2937;"
    for i, (_, r) in enumerate(df.iterrows(), start=1):
        code = code_of(r["Mã hàng"])
        price = float(r["Đơn giá (VND)"] or 0)
        name = clean(r["Diễn giải"]) if code else (clean(r["BOQ gốc"]) or clean(r["Diễn giải"]))
        flag = (not code) or price <= 0
        bg = "background:#FFF2A8;color:#111;" if flag else ""
        amount = float(r["SL"] or 0) * price
        cells = [(str(i), "center"), (code, "center"), (name, "left"), (clean(r["Xuất xứ"]), "center"),
                 (clean(r["ĐVT"]) or "Cái", "center"), (f"{float(r['SL'] or 0):,.2f}", "center"),
                 (_money(price), "center"), (_money(amount), "center"), (clean(r["Ghi chú"]), "left")]
        tds = "".join(f"<td style='{cell_style}{bg}text-align:{al}'>{esc(txt)}</td>" for txt, al in cells)
        rows.append(f"<tr>{tds}</tr>")
    th = "border:1px solid #9FB3C8;padding:6px;background:#E8F0F8;color:#1F4E79;font-weight:700;"
    td = "border:1px solid #9FB3C8;padding:5px;color:#1F2937;text-align:center;"
    head = (
        f"<table style='width:100%;border-collapse:collapse;margin-bottom:12px;font-size:14px;color:#1F2937'>"
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
        "<table style='width:100%;margin-top:8px;font-size:14px;color:#1F2937'>"
        f"<tr><td style='text-align:right;padding:4px'><b>Cộng:</b></td><td style='width:200px;text-align:right'>{_money(cong)}</td></tr>"
        f"<tr><td style='text-align:right;padding:4px'><b>Thuế GTGT (8%):</b></td><td style='text-align:right'>{_money(thue)}</td></tr>"
        f"<tr><td style='text-align:right;padding:4px;color:#1F4E79'><b>Tổng cộng (VND):</b></td>"
        f"<td style='text-align:right;color:#1F4E79;font-weight:700'>{_money(tong)}</td></tr></table>")
    foot = (
        "<div style='margin-top:12px;font-size:14px;color:#1F2937'>"
        f"<p><b>Lưu ý:</b> {esc(f['note'])}</p>"
        f"<p><b>Giao hàng:</b> {esc(f['deliv1'])}<br>{esc(f['deliv2'])}</p>"
        f"<p><b>Thanh toán:</b><br>{esc(f['payment']).replace(chr(10), '<br>')}</p></div>")
    return head + table + totals + foot


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
                     Paragraph(clean(r["Ghi chú"]), small)])
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


def label_item(it):
    return f"[{it['ma']}] {it['ten']} - {it['don_gia']:,.0f} đ"


def _fmt_date(d):
    """Nhận ngày dạng date hoặc chuỗi, trả về chuỗi dd/mm/yyyy."""
    return d.strftime("%d/%m/%Y") if hasattr(d, "strftime") else str(d)


def build_excel(data: pd.DataFrame, H: dict, F: dict, customer: bool = False) -> bytes:
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
            cell.value = None if (isinstance(val, str) and val == "") else val

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
        safe_write(6, 2, H['to'])
        safe_write(7, 2, H['add'])
        safe_write(8, 2, H['att'])
        safe_write(9, 2, H['mobile'])
        safe_write(10, 2, H['email'])
        safe_write(11, 2, H['ref'])
        safe_write(6, 7, H['no'])
        safe_write(7, 7, _fmt_date(H["date"]))
        safe_write(8, 7, H['revise'])
        safe_write(9, 7, H['validity'])
        safe_write(10, 7, H['project'])

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
        item_h = ws.row_dimensions[last_item_row].height or ws.row_dimensions[start_row].height or 43.9
        if n > available:
            extra = n - available
            shift_rows(total_row, extra)
            for r in range(total_row, total_row + extra):
                ws.row_dimensions[r].height = item_h  # dòng thêm cao bằng dòng mẫu
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

        # 4) Ghi hàng hóa (nền vùng hàng hóa về trắng như mẫu; chỉ tô vàng dòng cần soát ở bản nội bộ)
        flag_fill = PatternFill("solid", fgColor="FFF2A8")
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
            if clean(r["Ghi chú"]):
                notes.append(clean(r["Ghi chú"]))
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
                safe_write(r, 3, F['note'])
            elif k2.startswith("giao hang"):
                safe_write(r, 3, F['deliv1'])
                if clean(F.get('deliv2')) and not clean(ws.cell(row=r + 1, column=2).value) \
                        and not clean(ws.cell(row=r + 1, column=3).value):
                    safe_write(r + 1, 3, F['deliv2'])
            elif k3.startswith("nhap hang") or "giay luu thong" in k3:
                safe_write(r, 3, F['deliv2'])
            elif k2.startswith("thanh toan"):
                # Khối thanh toán của mẫu: gom hết các dòng trống phía dưới nhãn, đến dòng "Thanh toán bằng..."
                lines = [ln.strip() for ln in F['payment'].split("\n") if ln.strip()]
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

        buf = BytesIO()
        wb.save(buf)
        out = restore_template_drawing(clean_xlsx_for_excel(buf.getvalue()))
        import zipfile as _zf
        with _zf.ZipFile(BytesIO(out)) as _z:
            if _z.testzip() is not None:
                raise ValueError("File Excel tạo ra bị hỏng (zip lỗi).")
            names = _z.namelist()
            if len(names) != len(set(names)):
                raise ValueError("File Excel tạo ra bị trùng phần tử bên trong.")
        openpyxl.load_workbook(BytesIO(out))  # đọc lại thử: lỗi thì báo ngay, không đưa file hỏng cho người dùng
        return out




def strip_drawing(xlsx_bytes: bytes) -> bytes:
    """Bỏ hoàn toàn phần hình vẽ (logo, đường kẻ) khỏi file Excel để tránh lỗi mở file."""
    import re as _re
    import zipfile as _zf
    src = _zf.ZipFile(BytesIO(xlsx_bytes))
    out = BytesIO()
    with _zf.ZipFile(out, "w", _zf.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            name = item.filename
            if name.startswith("xl/drawings/") or name.startswith("xl/media/"):
                continue
            content = src.read(name)
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                content = _re.sub(r"<drawing\b[^>]*/>", "", content.decode("utf-8")).encode("utf-8")
            elif name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
                content = _re.sub(r'<Relationship [^>]*relationships/drawing"[^>]*/>', "",
                                  content.decode("utf-8")).encode("utf-8")
            elif name == "[Content_Types].xml":
                content = _re.sub(r'<Override PartName="/xl/drawings/[^"]*"[^>]*/>', "",
                                  content.decode("utf-8")).encode("utf-8")
            dst.writestr(item, content)
    return out.getvalue()
