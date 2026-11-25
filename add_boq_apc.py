"""
add_boq_apc.py - Gắn mã VAN APC cho các dòng BOQ đã biết.

- Mã và tên chuẩn lấy TRỰC TIẾP từ sheet 'VAN APC' của BangGia.xlsx (không đoán).
- Ghi vào seed_mappings.json (upsert theo BOQ) VÀ ghi đè learned_knowledge.json,
  vì KnowledgeStore ưu tiên mapping đã học hơn seed.
- Dòng nào không có mã trong bảng giá APC thì KHÔNG ghi, in ra để xử lý tay.
"""
import json, os, re, sys
import openpyxl

HERE = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
SRC = "BOQ khách - mã VAN APC (đối chiếu bảng giá)"

def norm(t):
    s = str(t or "").lower().replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()

# (tên BOQ, ĐVT, mã APC, tên chuẩn APC)
ROWS = [
    ("Van xả khí tự động kèm van", "Cái", "MH-AMARV", "Van xả khí tự động"),
    ("Van xả khí tự động buồng bơm", "Cái", "MH-AMARV", "Van xả khí tự động"),
    ("Van bi DN15", "Cái", "QF-300", "Van bi DN15"),
    ("Van bi DN25", "Cái", "QF-300", "Van bi DN25"),
    ("Van bướm có công tắc giám sát điện DN80", "Cái", "WD-371X", "Van bướm có giám sát DN80"),
    ("Van bướm có công tắc giám sát điện DN100", "Cái", "WD-371X", "Van bướm có giám sát DN100"),
    ("Van bướm có công tắc giám sát điện DN150", "Cái", "WD-371X", "Van bướm có giám sát DN150"),
    ("Van 1 chiều DN50", "Cái", "MX-XQH-350F", "Van 1 chiều DN50"),
    ("Van 1 chiều DN80", "Cái", "MX-XQH-350F", "Van 1 chiều DN80"),
    ("Van 1 chiều DN100", "Cái", "MX-XQH-350F", "Van 1 chiều DN100"),
    ("Van 1 chiều DN150", "Cái", "MX-XQH-350F", "Van 1 chiều DN150"),
    ("Khớp nối mềm DN50", "Cái", "RJ-01", "Khớp nối mềm DN50"),
    ("Khớp nối mềm DN80", "Cái", "RJ-01", "Khớp nối mềm DN80"),
    ("Khớp nối mềm DN100", "Cái", "RJ-01", "Khớp nối mềm DN100"),
    ("Công tắc dòng chảy DN80", "Cái", "ZSJZ-YM-450", "Công tắc dòng chảy"),
    ("Công tắc dòng chảy DN100", "Cái", "ZSJZ-YM-450", "Công tắc dòng chảy"),
    ("Van cổng O&Y DN25", "Cái", "GV", "Van cổng DN25"),
    ("Van cổng O&Y DN65", "Cái", "Z41-300", "Van Cổng OS&Y DN65"),
    ("Van cổng O&Y DN100", "Cái", "Z41-300", "Van Cổng OS&Y DN100"),
    ("Van cổng O&Y DN150", "Cái", "Z41-300", "Van Cổng OS&Y DN150"),
    ("Van cổng có công tắc giám sát điện DN80", "Cái", "Z41-300+ OYSS - 1", "Van Cổng OS&Y Kèm Giám Sát DN80"),
    ("Van cổng có công tắc giám sát điện DN100", "Cái", "Z41-300+ OYSS - 1", "Van Cổng OS&Y Kèm Giám Sát DN100"),
    ("Van cổng có công tắc giám sát điện DN150", "Cái", "Z41-300+ OYSS - 1", "Van Cổng OS&Y Kèm Giám Sát DN150"),
    ("Van giảm áp DN100", "Cái", "PRV", "Van Giảm Áp DN100"),
    ("Van giảm áp DN150", "Cái", "PRV", "Van Giảm Áp DN150"),
    ("Van bảo vệ vỏ bơm", "Cái", "55L-60", "Van bảo vệ vỏ bơm"),
    ("Y lọc DN50", "Cái", "YS-01", "Van Y Lọc DN50"),
    ("Y lọc DN100", "Cái", "YS-300-FF", "Van Y lọc DN100"),
    ("Van báo động DN100", "Bộ", "SSBJ-300F", "Van báo động DN100"),
    ("Van báo động DN150", "Bộ", "SSBJ-300F", "Van báo động DN150"),
]

wb = openpyxl.load_workbook(os.path.join(HERE, "BangGia.xlsx"), data_only=True)
apc = []
for r in wb["VAN APC"].iter_rows(min_row=2, values_only=True):
    ma, ten = (r[0] or "").strip(), (r[1] or "").strip()
    if ma and ten:
        apc.append({"ma": ma, "ten": re.sub(r"\s+", " ", ten),
                    "xuat_xu": (r[2] or "").replace("\n", "").strip() or "APC/China"})

def find_apc(ma, ten):
    m, t = norm(ma), norm(ten)
    for a in apc:
        if norm(a["ma"]) == m and norm(a["ten"]) == t:
            return a
    return None

seed_path = os.path.join(HERE, "seed_mappings.json")
learn_path = os.path.join(HERE, "learned_knowledge.json")
seed = json.load(open(seed_path, encoding="utf-8"))
learned = json.load(open(learn_path, encoding="utf-8"))
mappings = learned.setdefault("mappings", {})

existing = {norm(r["boq"]): i for i, r in enumerate(seed)}
added = updated = 0
ok, skipped = 0, []
for boq, dvt, ma, ten in ROWS:
    hit = find_apc(ma, ten)
    if not hit:
        skipped.append((boq, ma, ten))
        continue
    ok += 1
    rec = {"boq": boq, "ma": hit["ma"], "ten": hit["ten"], "xuat_xu": hit["xuat_xu"],
           "dvt": dvt, "don_gia": 0.0, "nguon": SRC}
    k = norm(boq)
    if k in existing:
        seed[existing[k]] = rec; updated += 1
    else:
        seed.append(rec); existing[k] = len(seed) - 1; added += 1
    mappings[k] = {"ma": rec["ma"], "ten": rec["ten"], "xuat_xu": rec["xuat_xu"], "dvt": dvt,
                   "don_gia": 0.0, "count": 5, "updated": "BOQ APC bổ sung", "nguon": SRC}

json.dump(seed, open(seed_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(learned, open(learn_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

print(f"seed: added={added} updated={updated} total={len(seed)}")
print(f"learned ghi đè: {ok}")
print("KHÔNG có mã trong bảng giá APC (chưa ghi):")
for boq, ma, ten in skipped:
    print(f"  - {boq}   [mã điền: {ma}]")
