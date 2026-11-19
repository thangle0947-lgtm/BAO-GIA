import json, re

def norm(t):
    s = str(t or "").lower().replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()

SRC = "BOQ khách - đầu phun (bổ sung thủ công)"
DOWN = {"ma": "T-ZSTX 80-68°C Q5", "ten": "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C",
        "xuat_xu": "Jilong/China", "dvt": "Cái"}
UP = {"ma": "T-ZSTZ 80-68°C Q5", "ten": "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C",
      "xuat_xu": "Jilong/China", "dvt": "Cái"}
NEW = [
    ('Đầu phun hướng xuống 1/2", K=5.6, 68 độ C', DOWN),
    ('Đầu phun hướng lên 1/2", K=5.6, 68 độ C', UP),
    ('Đầu phun khô hướng xuống 1/2", K=5.6, 68 độ C',
     {"ma": "T-ZSTGX 80-68°C Q5", "ten": "Đầu phun chữa cháy khô hướng xuống phản ứng tiêu chuẩn, K=5.6, DN25, 68°C",
      "xuat_xu": "Jilong/China", "dvt": "Cái"}),
    ('Đầu phun âm trần hướng xuống 1/2", K=5.6, 68 độ C dùng cho phòng sạch',
     {"ma": "ZSTDY 80-68°C G Q5", "ten": "Đầu phun chữa cháy âm trần phản ứng tiêu chuẩn, K=5.6, DN15, 68°C",
      "xuat_xu": "Jilong/China", "dvt": "Bộ"}),
    ('Đầu phun hướng xuống 1/2", K=5.6, 68 độ C, đầu phun dự phòng', DOWN),
    ('Đầu phun hướng lên 1/2", K=5.6, 68 độ C, đầu phun dự phòng', UP),
]

seed = json.load(open("seed_mappings.json", encoding="utf-8"))
existing = {norm(r["boq"]): i for i, r in enumerate(seed)}
added, updated = 0, 0
for boq, d in NEW:
    rec = {"boq": boq, "ma": d["ma"], "ten": d["ten"], "xuat_xu": d["xuat_xu"],
           "dvt": d["dvt"], "don_gia": 0.0, "nguon": SRC}
    k = norm(boq)
    if k in existing:
        seed[existing[k]] = rec   # cập nhật nếu đã có
        updated += 1
    else:
        seed.append(rec)
        existing[k] = len(seed) - 1
        added += 1

json.dump(seed, open("seed_mappings.json", "w", encoding="utf-8"), ensure_ascii=False, indent=2)
print(f"added={added} updated={updated} total={len(seed)}")

# Kiểm tra: không trùng khóa, mọi bản ghi có đủ trường, mã mới khớp mô tả
again = json.load(open("seed_mappings.json", encoding="utf-8"))
keys = [norm(r["boq"]) for r in again]
assert len(keys) == len(set(keys)), "trùng khóa"
assert all(r["ma"] and r["boq"] for r in again), "thiếu trường"
for boq, d in NEW:
    hit = [r for r in again if norm(r["boq"]) == norm(boq)][0]
    print(hit["ma"].ljust(20), "|", boq[:55])
