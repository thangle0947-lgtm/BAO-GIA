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


def item_key(item) -> str:
    """Khóa duy nhất của một mục bảng giá: mã + tên (vì một mã có thể dùng cho nhiều kích thước)."""
    return f"{item['ma']} | {item['ten']}"


def code_of(v) -> str:
    """Lấy phần mã từ khóa 'mã | tên'."""
    return clean(v).split(" | ")[0].strip()


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
    # Phản ứng nhanh: có "nhanh"/"quick", cụm RTI < 50 (m/s)^0.5 (VN/CN), hoặc ESFR
    if re.search(r"nhanh|quick|esfr|nhỏ hơn 50|热响应", t):
        response = {"nhanh"}
    elif "tiêu chuẩn" in t or "standard" in t:
        response = {"tiêu chuẩn"}
    elif re.search(r"đầu phun|sprinkler|喷淋", t):
        response = {"tiêu chuẩn"}  # đầu phun không ghi gì thì mặc định tiêu chuẩn
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


VALVE_FAMILIES = [
    ("pressure_relief", r"xả áp"),
    ("pressure_reduce", r"giảm áp"),
    ("air_release", r"xả khí"),
    ("safety", r"an toàn|relief valve"),
    ("casing_relief", r"bảo vệ vỏ bơm|casing relief|relif valve"),
    ("alarm", r"báo động|alarm"),
    ("gate", r"van cổng|gate valve|\bcổng\b"),
    ("butterfly", r"van bướm|butterfly"),
    ("check", r"1 chiều|một chiều|check valve"),
    ("float", r"phao|float valve"),
    ("strainer", r"rọ hút|y lọc|foot valve"),
    ("flow_switch", r"công tắc dòng chảy|flow switch"),
    ("pressure_switch", r"công tắc áp suất|pressure switch"),
    ("gauge", r"đồng hồ|pressure gauge|flow meter"),
    ("solenoid", r"điện từ"),
]


def families(text) -> set:
    t = norm(text)
    return {name for name, pat in VALVE_FAMILIES if re.search(pat, t)}


def gate_type(text) -> str:
    """Van cổng: 'nrs' = ty chìm, 'osy' = ty nổi / OS&Y, '' = không ghi."""
    t = norm(text)
    if re.search(r"nrs|ty chìm", t):
        return "nrs"
    if re.search(r"os&y|ty nổi", t):
        return "osy"
    return ""


def is_supervised(text) -> bool:
    return bool(re.search(r"giám sát|tín hiệu|oyss|bướm điện|supervis", norm(text)))


def items_conflict(query, cand) -> bool:
    """True nếu mục bảng giá không phù hợp với mô tả BOQ, dù tên có vài chữ giống nhau."""
    if specs_conflict(extract_specs(query), extract_specs(cand)):
        return True
    fq, fc = families(query), families(cand)
    if fq and not (fq & fc):
        return True  # mô tả nêu loại van cụ thể mà mục bảng giá không có loại đó
    if "gate" in fq and "gate" in fc:
        gq, gc = gate_type(query), gate_type(cand)
        if gq and gc and gq != gc:
            return True
        if not gq and gc == "nrs":
            return True  # không ghi ty chìm thì mặc định OS&Y, không dùng NRS
    if ("gate" in fq and "gate" in fc) or ("butterfly" in fq and "butterfly" in fc):
        if is_supervised(query) != is_supervised(cand):
            return True
    return False


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
def choice_items(price_items, store) -> list:
    """Danh sách mã cho ô chọn: mục bảng giá đang chọn + mã đã học/seed + mã nhập tay.
    Không ghi đè mục bảng giá; mã trùng (mã | tên) chỉ xuất hiện một lần."""
    out = list(price_items)
    seen = {item_key(p) for p in out if p.get("ma")}
    cands = list(store.data.get("codes", [])) + list(store.data.get("mappings", {}).values())
    for e in cands:
        if not e.get("ma") or not e.get("ten"):
            continue
        it = {"sheet": "ĐÃ HỌC", "ma": str(e["ma"]), "ten": str(e["ten"]), "xuat_xu": e.get("xuat_xu", "") or "",
              "dvt": e.get("dvt") or "Cái", "don_gia": float(e.get("don_gia") or 0)}
        k = item_key(it)
        if k not in seen:
            seen.add(k)
            out.append(it)
    return out


def load_price_items(path, sheets=None) -> list:
    return apply_overrides(_load_price_items_base(path, sheets))


def _load_price_items_base(path, sheets=None) -> list:
    items = _load_price_items_from_excel(path, sheets)
    # Mục bổ sung (mã chưa có trong bảng giá), luôn được thêm vào
    extra_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extra_items.json")
    if os.path.exists(extra_path):
        try:
            with open(extra_path, "r", encoding="utf-8") as f:
                existing = {item_key(i) for i in items}
                for e in json.load(f):
                    item = {"sheet": e.get("sheet", "BỔ SUNG"), "ten": e["ten"], "ma": e["ma"],
                            "xuat_xu": e.get("xuat_xu", ""), "dvt": e.get("dvt", "Cái"),
                            "don_gia": float(e.get("don_gia", 0))}
                    key = item_key(item)
                    if key not in existing:
                        items.append(item)
                    elif item["don_gia"] > 0:
                        # Giá hãng trong mục bổ sung ghi đè giá bảng giá cũ (0 hoặc lỗi thời)
                        for it in items:
                            if item_key(it) == key:
                                it["don_gia"] = item["don_gia"]
                                break
        except (json.JSONDecodeError, OSError, KeyError):
            pass
    return items


OVERRIDE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gia_ghi_de.json")


def load_overrides() -> dict:
    """Đơn giá do người dùng sửa trên web: {khóa mã|tên: đơn giá}."""
    if not os.path.exists(OVERRIDE_PATH):
        return {}
    try:
        with open(OVERRIDE_PATH, "r", encoding="utf-8") as f:
            return {k: float(v) for k, v in json.load(f).items()}
    except (json.JSONDecodeError, OSError, ValueError):
        return {}


def save_override(item, price: float):
    data = load_overrides()
    data[item_key(item)] = float(price)
    with open(OVERRIDE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def apply_overrides(items: list) -> list:
    ov = load_overrides()
    for it in items:
        k = item_key(it)
        if k in ov:
            it["don_gia"] = ov[k]
    return items


def _load_price_items_from_excel(path, sheets=None) -> list:
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


def find_layout(raw, layouts, min_sim=0.8):
    """Tìm cách đọc đã nhớ khớp với file này. Trả về (entry, chỉ số dòng tiêu đề) hoặc (None, None)."""
    best, best_row, best_sim = None, None, 0.0
    for entry in layouts:
        ref = set(entry["headers"])
        if not ref:
            continue
        for i in range(min(15, len(raw))):
            sig = {norm(v) for v in raw.iloc[i].tolist() if clean(v)}
            if not sig:
                continue
            sim = len(sig & ref) / len(sig | ref)
            if sim > best_sim:
                best, best_row, best_sim = entry, i, sim
    if best_sim >= min_sim:
        return best, best_row
    return None, None


def layout_columns(df, entry):
    cols = list(df.columns)
    return {k: (cols[p] if p is not None and p < len(cols) else None)
            for k, p in entry["positions"].items()}


def promote_header(raw: pd.DataFrame, header_row=None):
    """Dùng dòng tiêu đề làm tên cột. Trả về (df dữ liệu, chỉ số dòng tiêu đề)."""
    h = find_header_row(raw) if header_row is None else header_row
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
        self.data = {"mappings": {}, "price_history": {}, "corrections": [], "layouts": []}
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    self.data.update(json.load(f))
            except (json.JSONDecodeError, OSError):
                pass  # file hỏng: bắt đầu lại, không làm sập app
        self.data.setdefault("layouts", [])
        self.data.setdefault("codes", [])  # mã tự nhập, luôn hiện trong danh sách chọn
        # Nạp mẫu ban đầu (seed_mappings.json cạnh module), không ghi đè cái đã học
        seed_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "seed_mappings.json")
        if os.path.exists(seed_path):
            try:
                with open(seed_path, "r", encoding="utf-8") as f:
                    for e in json.load(f):
                        key = norm(e["boq"])
                        if key and key not in self.data["mappings"]:
                            self.data["mappings"][key] = {
                                "ma": e["ma"], "ten": e["ten"], "xuat_xu": e.get("xuat_xu", ""),
                                "dvt": e.get("dvt", ""), "don_gia": e.get("don_gia", 0),
                                "count": 1, "updated": "seed", "nguon": e.get("nguon", "seed")}
            except (json.JSONDecodeError, OSError, KeyError):
                pass
        # Dọn các ghép đã học bị lệch kích thước do lỗi cũ (cùng mã, khác DN)
        bad = [k for k, v in self.data["mappings"].items()
               if v.get("ten") and items_conflict(k, v["ten"])]
        for k in bad:
            del self.data["mappings"][k]

    def add_code(self, item: dict):
        """Lưu một mã nhập tay (không gắn với BOQ nào) để lần sau có trong danh sách chọn."""
        key = item_key(item)
        keep = [c for c in self.data.get("codes", []) if item_key(c) != key]
        keep.append({k: item.get(k, "") for k in ("ma", "ten", "xuat_xu", "dvt", "don_gia")})
        self.data["codes"] = keep
        self.save()

    def save_layout(self, raw, header_row, df_columns, cols):
        """Ghi nhớ cách chọn cột cho cấu trúc tiêu đề này (lưu theo vị trí cột)."""
        headers = sorted({norm(v) for v in raw.iloc[header_row].tolist() if clean(v)})
        cols_list = list(df_columns)
        positions = {k: (cols_list.index(v) if v in cols_list else None) for k, v in cols.items()}
        entry = {"headers": headers, "header_row": header_row, "positions": positions,
                 "updated": datetime.now().isoformat(timespec="seconds")}
        self.data["layouts"] = [l for l in self.data["layouts"] if l["headers"] != headers] + [entry]
        self.save()

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
        if not learned and store.data["mappings"]:
            key = norm(boq_text)
            best_key, best_r = None, 0.0
            for k in store.data["mappings"]:
                r = difflib.SequenceMatcher(None, key, k).ratio()
                if r > best_r:
                    best_key, best_r = k, r
            if best_r >= 0.9:
                learned = store.data["mappings"][best_key]
                if learned.get("ten") and items_conflict(boq_text, learned["ten"]):
                    learned = None
        if learned and learned.get("ma"):
            match = next((p for p in price_items
                          if p["ma"] == learned["ma"] and p["ten"] == learned.get("ten")), None)
            # Chỉ dùng mục đã học nếu kích thước/thông số không mâu thuẫn với BOQ
            if match and items_conflict(boq_text, match["ten"]):
                match = None
            if match:
                suggestions.append({"item": match, "score": 1.0, "source": "learned",
                                    "reason": f"Đã ghép tay {learned['count']} lần trước đây"})

    q_tok = tokens(boq_text)
    q_spec = extract_specs(boq_text)
    scored = []
    for item in price_items:
        if not item["ma"]:
            continue  # dòng tiêu đề nhóm trong bảng giá
        if any(item_key(s["item"]) == item_key(item) for s in suggestions):
            continue
        if items_conflict(boq_text, item["ten"]):
            continue
        p_tok = tokens(item["ten"])
        if not q_tok or not p_tok:
            continue
        jacc = len(q_tok & p_tok) / len(q_tok | p_tok)
        ratio = difflib.SequenceMatcher(None, norm(boq_text), norm(item["ten"])).ratio()
        score = 0.6 * jacc + 0.4 * ratio
        if is_supervised(boq_text) and re.search(r"tín hiệu|kèm giám sát", norm(item["ten"])):
            score += 0.25  # van có giám sát: ưu tiên mã tín hiệu / OS&Y kèm giám sát đúng quy định
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
def parse_qty(v):
    """Đọc số lượng từ ô Excel: hiểu '1,081.00', '1.081', '417', '-' (không có SL -> NaN)."""
    if v is None:
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\u00a0", "").replace(" ", "")
    if s in ("", "-", "–", "—", "None", "nan", "NaN"):
        return float("nan")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):   # 1.081,00 (kiểu VN)
            s = s.replace(".", "").replace(",", ".")
        else:                             # 1,081.00 (kiểu US)
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    return pd.to_numeric(s, errors="coerce")


def merge_duplicate_rows(rows):
    """Gộp các dòng có cùng BOQ gốc (không phân biệt hoa thường, khoảng trắng), cộng SL.
    Giữ thứ tự xuất hiện đầu tiên; các trường khác lấy từ dòng đầu."""
    merged, index = [], {}
    for r in rows:
        key = norm(r.get("BOQ gốc"))
        if key in index:
            m = merged[index[key]]
            m["SL"] = float(m.get("SL") or 0) + float(r.get("SL") or 0)
            m["Duyệt"] = bool(m.get("Duyệt")) and bool(r.get("Duyệt"))
        else:
            index[key] = len(merged)
            merged.append(dict(r))
    return merged


def build_quote_rows(df_raw, cols, price_items, store, use_llm=False):
    """Tạo dòng báo giá. AI chỉ ĐỀ XUẤT mã (cột Mã đề xuất). Mã chỉ được điền
    vào Mã hàng khi người dùng duyệt, trừ các ghép đã học và đã duyệt trước đó."""
    rows = []
    stats = {"total": 0, "proposed": 0, "learned": 0, "review": 0, "none": 0}
    for _, src in df_raw.iterrows():
        name = clean(src.get(cols["name"]))
        if not name:
            continue
        qty_raw = parse_qty(src.get(cols["qty"]))
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
            alt = [x for x in sugg[:3] if proposal is None or item_key(x["item"]) != item_key(proposal["item"])]
            alternatives = " | ".join(f"[{x['item']['ma']}] {x['item']['ten'][:40]}" for x in alt[:2])
        if proposal is not None and proposal["score"] < AUTO_FILL and source == "AI":
            stats["review"] += 1

        default_unit = "Mét" if "cáp" in name.lower() else "Cái"
        row = {"Xóa": False,
               "Duyệt": source == "Đã học",
               "BOQ gốc": full_text,
               "Mã đề xuất": item_key(proposal["item"]) if proposal else "",
               "Độ khớp": proposal["score"] if proposal else 0.0,
               "Nguồn": source,
               "Mã hàng": "", "Diễn giải": name, "Xuất xứ": "",
               "ĐVT": boq_unit or default_unit, "SL": qty,
               "Đơn giá (VND)": 0.0, "Thành tiền (VND)": 0.0,
               "Ghi chú": "",  # ghi chú nội bộ (đã học/ghép) nằm ở cột Nguồn, không đi vào file
               "Gợi ý": (f"Phương án khác: {alternatives}" if alternatives else
                         ("" if proposal else "⚠ Không có ứng viên, hãy ghép tay"))}
        rows.append(row)
    before = len(rows)
    rows = merge_duplicate_rows(rows)
    stats["merged"] = before - len(rows)
    return rows, stats


def fill_item(df, idx, item):
    df.at[idx, "Mã hàng"] = item_key(item)
    df.at[idx, "Diễn giải"] = item["ten"]
    df.at[idx, "Xuất xứ"] = item["xuat_xu"]
    df.at[idx, "ĐVT"] = item["dvt"] or df.at[idx, "ĐVT"]
    df.at[idx, "Đơn giá (VND)"] = item["don_gia"]


def sync_codes(df, edited_rows, price_items, store, seen: set) -> int:
    """Người dùng chọn mã trong danh sách: điền tên, giá, xuất xứ và ghi nhớ cặp BOQ -> mã."""
    by_key = {item_key(p): p for p in price_items if p["ma"]}
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
        item = by_key.get(code)
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


def apply_cell_edits(df, edited_rows):
    """Áp các ô người dùng vừa sửa trong bảng vào df (giữ nguyên chỉ số dòng)."""
    df = df.copy()
    for ridx, changes in (edited_rows or {}).items():
        try:
            ridx = int(ridx)
        except (TypeError, ValueError):
            continue
        if ridx not in df.index:
            continue
        for col, val in changes.items():
            if col in df.columns:
                df.at[ridx, col] = val
    return df


def remove_and_add_rows(df, state):
    """Xóa dòng đã tick 'Xóa' hoặc bị xóa trong bảng; thêm dòng mới. Trả về (df, có đổi cấu trúc)."""
    edited_rows = state.get("edited_rows") or {}
    to_drop = {int(r) for r, ch in edited_rows.items() if ch.get("Xóa") is True}
    to_drop |= {int(i) for i in (state.get("deleted_rows") or [])}
    added = state.get("added_rows") or []
    structural = bool(to_drop or added)
    df = df.drop(index=[i for i in to_drop if i in df.index])
    if added:
        df = pd.concat([df, pd.DataFrame(added)], ignore_index=True)
    return df.reset_index(drop=True), structural


def apply_approved(df, price_items, store, seen: set) -> int:
    """Điền mã đề xuất vào Mã hàng cho các dòng người dùng đã tích Duyệt.
    Khi duyệt ghép mới (không phải 'Đã học'), ghi nhớ để lần sau tự đề xuất đúng hơn."""
    by_key = {item_key(p): p for p in price_items if p["ma"]}
    applied = 0
    for idx in df.index:
        if not bool(df.at[idx, "Duyệt"]):
            continue
        code = clean(df.at[idx, "Mã đề xuất"])
        item = by_key.get(code)
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
        sl, gia, ma = r.get("sl") or 0, r.get("don_gia") or 0, code_of(r.get("ma"))
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
