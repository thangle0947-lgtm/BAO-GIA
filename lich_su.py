"""
lich_su.py - Lưu lịch sử báo giá đã xuất: thông tin để tra cứu + file Excel/PDF để mở lại.
Dữ liệu nằm trong thư mục lich_su/ cạnh app.py.
"""
import json
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_DIR = os.path.join(BASE_DIR, "lich_su")
INDEX_PATH = os.path.join(HISTORY_DIR, "lich_su.json")


def _load_index() -> list:
    if not os.path.exists(INDEX_PATH):
        return []
    try:
        with open(INDEX_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def _save_index(items: list):
    os.makedirs(HISTORY_DIR, exist_ok=True)
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def save_quote(info: dict, xlsx_bytes: bytes = None, pdf_bytes: bytes = None) -> dict:
    """info: so, khach, du_an, tong, so_dong, ban_gui_khach. Trả về bản ghi đã lưu."""
    items = _load_index()
    now = datetime.now()
    rec_id = now.strftime("%Y%m%d%H%M%S%f")
    folder = os.path.join(HISTORY_DIR, rec_id)
    os.makedirs(folder, exist_ok=True)
    rec = dict(info, id=rec_id, ngay=now.strftime("%d/%m/%Y %H:%M"), file_xlsx=None, file_pdf=None)
    if xlsx_bytes:
        rec["file_xlsx"] = os.path.join(folder, "bao_gia.xlsx")
        with open(rec["file_xlsx"], "wb") as f:
            f.write(xlsx_bytes)
    if pdf_bytes:
        rec["file_pdf"] = os.path.join(folder, "bao_gia.pdf")
        with open(rec["file_pdf"], "wb") as f:
            f.write(pdf_bytes)
    items.insert(0, rec)
    _save_index(items)
    return rec


def list_quotes(keyword: str = "") -> list:
    kw = (keyword or "").strip().lower()
    items = _load_index()
    if not kw:
        return items
    return [r for r in items if kw in " ".join(str(v) for v in r.values()).lower()]


def read_file(path: str):
    if path and os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    return None
