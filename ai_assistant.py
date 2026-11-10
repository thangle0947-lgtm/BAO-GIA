"""
ai_assistant.py - Lớp trợ lý AI cho hệ thống báo giá QN TECH.

- Đọc đúng cột BOQ.
- Gợi ý mã hàng từ nhiều sheet bảng giá, có ràng buộc DN / K / nhiệt độ / số loop.
- Học từ các lần người dùng ghép mã tay (lưu learned_knowledge.json).
- Kiểm tra báo giá trước khi xuất.
- (Tùy chọn) Claude API cho các dòng khó, cần biến môi trường ANTHROPIC_API_KEY.
"""
import json
import os
import re
import difflib
from datetime import datetime

import pandas as pd

LLM_MODEL = os.environ.get("QNTECH_LLM_MODEL", "claude-haiku-5-5")
AUTO_FILL = 0.55  # điểm tối thiểu để tự điền mã mà không cần người xem lại

COLUMNS = ["Xóa", "Duyệt", "BOQ gốc", "Mã đề xuất", "Độ khớp", "Nguồn",
           "Mã hàng", "Diễn giải", "Xuất xứ", "ĐVT", "SL",
           "Đơn giá (VND)", "Thành tiền (VND)", "Ghi chú", "Gợi ý"]


# ----------------------------------------------------------------- Tiện ích
def clean(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def clean_unit(v) -> str:
    """BOQ có ĐVT song ngữ dạng '项\nGói' -> lấy phần tiếng Việt (dòng cuối)."""
    s = clean(v)
    return s.split("\n")[-1].strip() if s else ""


def norm(text) -> str:
    s = clean(text).lower().replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()


def tokens(text) -> set:
    return set(re.findall(r"\w+", norm(text)))


def _num(x) -> str:
    """'8' và '8.0' cùng là '8'; '11,2' là '11.2'."""
    return "%g" % float(x.replace(",", "."))


def extract_specs(text) -> dict:
    t = norm(text)
    # Hướng đầu phun: kiểm tra "xuống" trước vì tên GS/Jilong viết "xuống lên"
    if "xuống" in t or "pendent" in t or "gắn trần" in t:
        direction = {"xuống"}
    elif "hướng lên" in t or "upright" in t:
        direction = {"lên"}
    else:
        direction = set()
    if "nhanh" in t or "quick" in t:
        response = {"nhanh"}
    elif "tiêu chuẩn" in t or "standard" in t:
        response = {"tiêu chuẩn"}
    else:
        response = set()
    return {
        "dn": set(re.findall(r"(?:dn|phi|ø|ф)\s*(\d+)", t)),
        "k": set(_num(x) for x in re.findall(r"\bk\s*=?\s*(\d+[.,]?\d*)", t)),
        "temp": set(_num(x) for x in re.findall(r"(\d+)\s*(?:độ|°c|c\b)", t)),
        "loop": set(re.findall(r"(\d+)\s*loop", t)),
        "dir": direction,
        "resp": response,
    }


def specs_conflict(a: dict, b: dict) -> bool:
    """Hai bên cùng có một loại thông số nhưng không có giá trị chung -> không khớp."""
    return any(a[k] and b[k] and not (a[k] & b[k]) for k in a)


def has_llm() -> bool:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return False
    try:
        import anthropic  # noqa: F401
        return True
    except ImportError:
        return False


# ----------------------------------------------------------- Đọc bảng giá
def load_price_items(path, sheets=None) -> list:
    items = []
    if not path or not os.path.exists(path):
        return items
    xl = pd.ExcelFile(path)
    for sh in (sheets or xl.sheet_names):
        if sh not in xl.sheet_names:
            continue
        df = pd.read_excel(path, sheet_name=sh).dropna(how="all")
        if "Tên chuẩn" not in df.columns:
            continue
        for _, r in df.iterrows():
            name = clean(r.get("Tên chuẩn"))
            if not name:
                continue
            price = pd.to_numeric(r.get("Đơn giá"), errors="coerce")
            items.append({
                "sheet": sh,
                "ten": name,
                "ma": clean(r.get("Mã SP")),
                "xuat_xu": clean(r.get("XUẤT XỨ")),
                "dvt": clean(r.get("ĐVT")),
                "don_gia": 0.0 if pd.isna(price) else float(price),
            })
    return items


# ------------------------------------------------------- Nhận diện cột BOQ
NAME_PATTERNS = [r"tên dự án", r"tên thiết bị", r"^tên\b", r"tên hàng", r"hạng mục",
                 r"mặt hàng", r"vật tư", r"diễn giải", r"mô tả", r"danh mục", r"thiết bị"]
QTY_PATTERNS = [r"tổng khối lượng", r"số lượng", r"\bsl\b", r"\bqty\b", r"\bkl\b", r"khối lượng"]
UNIT_PATTERNS = [r"đơn vị", r"\bđvt\b", r"\bdvt\b", r"^đv\b"]
SPEC_PATTERNS = [r"thông số", r"đặc tính", r"quy cách"]
PRICE_PATTERNS = [r"đơn giá", r"\bgiá\b"]


def _find(columns, patterns):
    for pat in patterns:
        for col in columns:
            if re.search(pat, norm(col)):
                return col
    return None


def find_header_row(raw: pd.DataFrame, max_scan=15) -> int:
    """Tìm dòng tiêu đề: dòng nào có cả ô tên lẫn ô khối lượng."""
    for i in range(min(max_scan, len(raw))):
        cells = [norm(v) for v in raw.iloc[i].tolist()]
        if _find(cells, NAME_PATTERNS) and _find(cells, QTY_PATTERNS):
            return i
    return 0


def promote_header(raw: pd.DataFrame):
    """Dùng dòng tiêu đề tìm được làm tên cột. Trả về (df dữ liệu, chỉ số dòng tiêu đề)."""
    h = find_header_row(raw)
    names, seen = [], {}
    for j, v in enumerate(raw.iloc[h].tolist()):
        name = clean(v) or f"cot_{j}"
        seen[name] = seen.get(name, 0) + 1
        names.append(name if seen[name] == 1 else f"{name}_{seen[name]}")
    data = raw.iloc[h + 1:].copy()
    data.columns = names
    return data.reset_index(drop=True), h


def _col_profile(series: pd.Series):
    """Đặc điểm nội dung của một cột: tỷ lệ số, độ dài chữ, có phải STT 1..n không."""
    vals = [clean(v) for v in series.tolist()]
    vals = [v for v in vals if v]
    if not vals:
        return None
    n = len(vals)
    nums = pd.to_numeric(pd.Series(vals), errors="coerce")
    numeric_ratio = float(nums.notna().mean())
    is_seq = numeric_ratio == 1 and [float(x) for x in nums] == [float(i) for i in range(1, n + 1)]
    return {
        "numeric": numeric_ratio,
        "avg_len": sum(len(v) for v in vals) / n,
        "distinct": len(set(vals)) / n,
        "seq": is_seq,
        "positive": float((nums.dropna() > 0).mean()) if numeric_ratio > 0 else 0.0,
        "n": n,
    }


def _guess_by_content(df, cols, profiles, key, used):
    """Đoán cột theo nội dung khi tiêu đề không khớp."""
    best, best_score = None, 0.0
    for c in cols:
        p = profiles.get(c)
        if c in used or p is None:
            continue
        if key == "qty":
            # Số, dương, không phải STT 1..n
            if p["numeric"] >= 0.8 and not p["seq"] and p["positive"] >= 0.8:
                score = p["numeric"] + p["positive"] + min(p["n"], 50) / 100
            else:
                continue
        elif key == "name":
            if p["numeric"] > 0.2:
                continue
            score = min(p["avg_len"], 60) / 60 * (0.5 + 0.5 * p["distinct"])
        elif key == "unit":
            if p["numeric"] > 0.2 or p["avg_len"] > 12 or (p["distinct"] > 0.5 and p["n"] > 10):
                continue
            score = 1 - p["avg_len"] / 12
        else:
            continue
        if score > best_score:
            best, best_score = c, score
    return best


def detect_boq_columns(df: pd.DataFrame) -> dict:
    """Ưu tiên tiêu đề cột; nếu tiêu đề không khớp thì đoán theo nội dung."""
    cols = list(df.columns)
    profiles = {c: _col_profile(df[c]) for c in cols}
    result = {}

    result["qty"] = _find(cols, QTY_PATTERNS)
    result["price"] = _find(cols, PRICE_PATTERNS)
    used = {result["qty"], result["price"]} - {None}

    result["name"] = _find(cols, NAME_PATTERNS)
    if result["name"] is None:
        result["name"] = _guess_by_content(df, cols, profiles, "name", used)
    used.add(result["name"])

    if result["qty"] is None:
        result["qty"] = _guess_by_content(df, cols, profiles, "qty", used)
    used.add(result["qty"])

    result["unit"] = _find(cols, UNIT_PATTERNS)
    if result["unit"] is None:
        result["unit"] = _guess_by_content(df, cols, profiles, "unit", used)
    used.add(result["unit"])

    result["spec"] = _find([c for c in cols if c not in used], SPEC_PATTERNS)
    return result


# --------------------------------------------------------- Bộ nhớ học hỏi
class KnowledgeStore:
    def __init__(self, path="learned_knowledge.json"):
        self.path = path
        self.data = {"mappings": {}, "price_history": {}, "corrections": []}
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    self.data.update(json.load(f))
            except (json.JSONDecodeError, OSError):
                pass  # file hỏng: bắt đầu lại, không làm sập app

    def save(self):
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get_mapping(self, boq_text):
        return self.data["mappings"].get(norm(boq_text))

    def learn_mapping(self, boq_text, item: dict):
        key = norm(boq_text)
        if not key:
            return
        prev = self.data["mappings"].get(key)
        count = prev["count"] + 1 if prev and prev.get("ma") == item["ma"] else 1
        self.data["mappings"][key] = {
            "ma": item["ma"], "ten": item["ten"], "xuat_xu": item["xuat_xu"],
            "dvt": item["dvt"], "don_gia": item["don_gia"],
            "count": count, "updated": datetime.now().isoformat(timespec="seconds"),
        }
        self.save()

    def log_correction(self, boq_text, before, after):
        self.data["corrections"].append({
            "time": datetime.now().isoformat(timespec="seconds"),
            "boq": norm(boq_text), "before": before, "after": after,
        })
        self.data["corrections"] = self.data["corrections"][-500:]
        self.save()

    def record_quote_prices(self, quote_rows):
        for r in quote_rows:
            if r.get("ma") and r.get("don_gia"):
                hist = self.data["price_history"].setdefault(r["ma"], [])
                hist.append({"gia": r["don_gia"], "ngay": datetime.now().date().isoformat()})
                self.data["price_history"][r["ma"]] = hist[-20:]
        self.save()


# ------------------------------------------------------------ Gợi ý ghép mã
def suggest_matches(boq_text, price_items, store: KnowledgeStore = None, top_k=3):
    """Trả về [{item, score, reason, source}], 'learned' là ghép đã học."""
    suggestions = []
    if store:
        learned = store.get_mapping(boq_text)
        if learned and learned.get("ma"):
            match = next((p for p in price_items if p["ma"] == learned["ma"]), None)
            if match:
                suggestions.append({"item": match, "score": 1.0, "source": "learned",
                                    "reason": f"Đã ghép tay {learned['count']} lần trước đây"})

    q_tok = tokens(boq_text)
    q_spec = extract_specs(boq_text)
    scored = []
    for item in price_items:
        if not item["ma"]:
            continue  # dòng tiêu đề nhóm trong bảng giá
        if any(s["item"]["ma"] == item["ma"] for s in suggestions):
            continue
        if specs_conflict(q_spec, extract_specs(item["ten"])):
            continue
        p_tok = tokens(item["ten"])
        if not q_tok or not p_tok:
            continue
        jacc = len(q_tok & p_tok) / len(q_tok | p_tok)
        ratio = difflib.SequenceMatcher(None, norm(boq_text), norm(item["ten"])).ratio()
        score = 0.6 * jacc + 0.4 * ratio
        if score >= 0.35:
            scored.append((score, item))

    scored.sort(key=lambda x: x[0], reverse=True)
    for score, item in scored[:max(0, top_k - len(suggestions))]:
        suggestions.append({"item": item, "score": round(score, 2), "source": "rule",
                            "reason": "Giống về tên và thông số"})
    return suggestions


def ask_llm_for_match(boq_text, candidates):
    """Trả về chỉ số ứng viên được chọn, hoặc None."""
    if not has_llm() or not candidates:
        return None
    try:
        import anthropic
        options = "\n".join(
            f"{i}. [{c['item']['ma']}] {c['item']['ten']} | {c['item']['don_gia']:,.0f} đ"
            for i, c in enumerate(candidates))
        prompt = (f"Hạng mục BOQ: \"{boq_text}\"\nCác mã ứng viên:\n{options}\n"
                  "Chọn ứng viên đúng nhất về loại thiết bị, DN, K, nhiệt độ, phản ứng. "
                  "Chỉ trả về một số thứ tự, hoặc -1 nếu không cái nào phù hợp.")
        client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
        resp = client.messages.create(model=LLM_MODEL, max_tokens=10,
                                      messages=[{"role": "user", "content": prompt}])
        m = re.search(r"-?\d+", resp.content[0].text)
        if not m:
            return None
        k = int(m.group())
        return k if 0 <= k < len(candidates) else None
    except Exception:
        return None


# ------------------------------------------------- Tạo dòng báo giá từ BOQ
def build_quote_rows(df_raw, cols, price_items, store, use_llm=False):
    """Tạo dòng báo giá. AI chỉ ĐỀ XUẤT mã (cột Mã đề xuất). Mã chỉ được điền
    vào Mã hàng khi người dùng duyệt, trừ các ghép đã học và đã duyệt trước đó."""
    rows = []
    stats = {"total": 0, "proposed": 0, "learned": 0, "review": 0, "none": 0}
    for _, src in df_raw.iterrows():
        name = clean(src.get(cols["name"]))
        if not name:
            continue
        qty_raw = pd.to_numeric(src.get(cols["qty"]), errors="coerce")
        boq_unit = clean_unit(src.get(cols["unit"])) if cols.get("unit") else ""
        if pd.isna(qty_raw) and not boq_unit:
            continue  # dòng tiêu đề mục như "I. Đầu phun chữa cháy"
        qty = 0.0 if pd.isna(qty_raw) else float(qty_raw)
        spec = clean(src.get(cols["spec"])) if cols.get("spec") else ""
        full_text = f"{name} {spec}".strip() if spec else name
        stats["total"] += 1

        sugg = suggest_matches(full_text, price_items, store, top_k=5)
        proposal, source, note = None, "", ""
        if sugg and sugg[0]["source"] == "learned":
            proposal, source = sugg[0], "Đã học"
            note = sugg[0]["reason"]
            stats["learned"] += 1
        elif sugg:
            k = ask_llm_for_match(full_text, sugg) if use_llm else None
            if k is not None:
                proposal, source, note = sugg[k], "Claude", "Claude đề xuất"
            else:
                proposal, source = sugg[0], "AI"
                note = f"Độ khớp {sugg[0]['score']}"
            stats["proposed"] += 1
        else:
            stats["none"] += 1

        alternatives = ""
        if sugg:
            alt = [x for x in sugg[:3] if proposal is None or x["item"]["ma"] != proposal["item"]["ma"]]
            alternatives = " | ".join(f"[{x['item']['ma']}] {x['item']['ten'][:40]}" for x in alt[:2])
        if proposal is not None and proposal["score"] < AUTO_FILL and source == "AI":
            stats["review"] += 1

        default_unit = "Mét" if "cáp" in name.lower() else "Cái"
        row = {"Xóa": False,
               "Duyệt": source == "Đã học",
               "BOQ gốc": full_text,
               "Mã đề xuất": proposal["item"]["ma"] if proposal else "",
               "Độ khớp": proposal["score"] if proposal else 0.0,
               "Nguồn": source,
               "Mã hàng": "", "Diễn giải": name, "Xuất xứ": "",
               "ĐVT": boq_unit or default_unit, "SL": qty,
               "Đơn giá (VND)": 0.0, "Thành tiền (VND)": 0.0,
               "Ghi chú": note,
               "Gợi ý": (f"Phương án khác: {alternatives}" if alternatives else
                         ("" if proposal else "⚠ Không có ứng viên, hãy ghép tay"))}
        rows.append(row)
    return rows, stats


def fill_item(df, idx, item):
    df.at[idx, "Mã hàng"] = item["ma"]
    df.at[idx, "Diễn giải"] = item["ten"]
    df.at[idx, "Xuất xứ"] = item["xuat_xu"]
    df.at[idx, "ĐVT"] = item["dvt"] or df.at[idx, "ĐVT"]
    df.at[idx, "Đơn giá (VND)"] = item["don_gia"]


def sync_codes(df, edited_rows, price_items, store, seen: set) -> int:
    """Người dùng chọn mã trong danh sách: điền tên, giá, xuất xứ và ghi nhớ cặp BOQ -> mã."""
    by_code = {}
    for p in price_items:
        if p["ma"]:
            by_code.setdefault(p["ma"], p)
    n = 0
    for ridx, changes in (edited_rows or {}).items():
        if "Mã hàng" not in changes:
            continue
        try:
            ridx = int(ridx)
        except (TypeError, ValueError):
            continue
        if ridx not in df.index:
            continue
        code = clean(df.at[ridx, "Mã hàng"])
        item = by_code.get(code)
        if not item:
            continue
        key = ("ma", ridx, code)
        if key in seen:
            continue
        seen.add(key)
        fill_item(df, ridx, item)
        df.at[ridx, "Nguồn"] = "Chọn tay"
        boq = clean(df.at[ridx, "BOQ gốc"])
        if boq:
            store.learn_mapping(boq, item)
            store.log_correction(boq, "", code)
        n += 1
    return n


def apply_approved(df, price_items, store, seen: set) -> int:
    """Điền mã đề xuất vào Mã hàng cho các dòng người dùng đã tích Duyệt.
    Khi duyệt ghép mới (không phải 'Đã học'), ghi nhớ để lần sau tự đề xuất đúng hơn."""
    by_code = {}
    for p in price_items:
        if p["ma"]:
            by_code.setdefault(p["ma"], p)
    applied = 0
    for idx in df.index:
        if not bool(df.at[idx, "Duyệt"]):
            continue
        code = clean(df.at[idx, "Mã đề xuất"])
        item = by_code.get(code)
        if not item:
            continue
        key = (int(idx), code)
        if key in seen:
            continue
        seen.add(key)
        current = clean(df.at[idx, "Mã hàng"])
        if current and current != code:
            continue  # người dùng đã chọn mã khác, không ghi đè
        fill_item(df, idx, item)
        df.at[idx, "Nguồn"] = df.at[idx, "Nguồn"] if df.at[idx, "Nguồn"] == "Đã học" else "Đã duyệt"
        boq = clean(df.at[idx, "BOQ gốc"])
        if boq and df.at[idx, "Nguồn"] != "Đã học":
            store.learn_mapping(boq, item)
            store.log_correction(boq, "", item["ma"])
        applied += 1
    return applied


# --------------------------------------------- Học từ việc người dùng sửa
def learn_from_editor_changes(edited_rows, df, price_items, store, seen: set) -> int:
    """edited_rows: st.session_state[key]['edited_rows'] của data_editor.
    df: dữ liệu TRƯỚC khi áp dụng chỉnh sửa. seen: tập (BOQ, mã) đã học trong phiên."""
    items_by_code = {p["ma"]: p for p in price_items if p["ma"]}
    learned = 0
    for ridx, changes in (edited_rows or {}).items():
        if "Mã hàng" not in changes:
            continue
        try:
            ridx = int(ridx)
        except (TypeError, ValueError):
            continue
        if ridx not in df.index:
            continue
        new_ma = clean(changes["Mã hàng"])
        if new_ma not in items_by_code:
            continue
        boq = clean(df.at[ridx, "BOQ gốc"]) or clean(df.at[ridx, "Diễn giải"])
        key = (norm(boq), new_ma)
        if not boq or key in seen:
            continue
        before = clean(df.at[ridx, "Mã hàng"])
        store.learn_mapping(boq, items_by_code[new_ma])
        store.log_correction(boq, before, new_ma)
        seen.add(key)
        learned += 1
    return learned


# ------------------------------------------------------------- Kiểm tra
def audit_quote(rows, store: KnowledgeStore = None) -> list:
    """rows: list dict có: ma, sl, don_gia. Trả về [{level, row, message}]."""
    issues = []
    seen_codes = {}
    stt_like = sum(1 for i, r in enumerate(rows, 1)
                   if r.get("sl") and float(r["sl"]) == i and float(r["sl"]).is_integer())
    if rows and stt_like >= max(3, 0.8 * len(rows)):
        issues.append({"level": "error", "row": 1,
                       "message": "Cột SL trùng gần như hoàn toàn số thứ tự, có thể đọc nhầm cột khối lượng"})
    for i, r in enumerate(rows, start=1):
        label = f"Dòng {i}"
        sl, gia, ma = r.get("sl") or 0, r.get("don_gia") or 0, clean(r.get("ma"))
        if not ma:
            issues.append({"level": "warn", "row": i, "message": f"{label}: chưa ghép mã hàng"})
        if not gia:
            issues.append({"level": "error", "row": i, "message": f"{label}: đơn giá = 0"})
        if not sl:
            issues.append({"level": "warn", "row": i, "message": f"{label}: số lượng = 0"})
        if ma:
            seen_codes.setdefault(ma, []).append(i)
        if store and ma and gia:
            hist = store.data["price_history"].get(ma, [])
            if hist:
                last = hist[-1]["gia"]
                if last and abs(gia - last) / last > 0.2:
                    issues.append({"level": "info", "row": i,
                                   "message": f"{label}: giá {gia:,.0f} lệch >20% so với lần báo trước ({last:,.0f})"})
    for code, rr in seen_codes.items():
        if len(rr) > 1:
            issues.append({"level": "info", "row": rr[0],
                           "message": f"Mã {code} xuất hiện ở dòng {', '.join(map(str, rr))}, kiểm tra có trùng không"})
    return issues
