"""Cập nhật mã đầu phun GS + nắp chụp / rọ bảo vệ vào seed_mappings.json và learned_knowledge.json.
Chạy trong thư mục chứa app: python update_gs_sprinklers.py"""
import json, os, re, sys

SRC_DIR = "/mnt/user-data/uploads"
OUT_DIR = "/mnt/user-data/outputs"
NGUON = "BOQ GS - mẫu người dùng"

def norm(t):
    s = str(t or "").lower().replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip()

# (BOQ text, mã GS, tên theo bảng giá, ĐVT, xuất xứ)
GS = [
    ("Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "GS251",
     "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun chữa cháy hướng xuống phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "GS252",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun chữa cháy xuống lên phản ứng nhanh, K=5.6, DN15, 68°C", "GS254",
     "Đầu phun chữa cháy xuống lên phản ứng nhanh, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "GS281",
     "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun chữa cháy hướng xuống phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "GS282",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun chữa cháy hướng xuống phản ứng tiêu chuẩn, K=11.2, DN20, 68°C", "GS212",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=11.2, DN20, 68°C", "Cái", "GS Korea/VN"),
    # Cách viết ngắn trong BOQ khách
    ("Đầu phun hướng lên K=5.6, 68°C", "GS251",
     "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng lên K=8, 68°C", "GS281",
     "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng xuống K=5.6, 68°C", "GS252",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng xuống phản ứng nhanh K=5.6, 68°C", "GS254",
     "Đầu phun chữa cháy xuống lên phản ứng nhanh, K=5.6, DN15, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng xuống K=11.2, 68°C", "GS212",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=11.2, DN20, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng xuống K=8, 68°C (Có rọ bảo vệ)", "GS282",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=8.0, DN20, 68°C", "Cái", "GS Korea/VN"),
    ("Đầu phun hướng xuống K=11.2, 68°C (Có rọ bảo vệ)", "GS212",
     "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=11.2, DN20, 68°C", "Cái", "GS Korea/VN"),
]

# Phụ kiện đi kèm: chưa có mã/giá trong bảng giá -> ma để trống, đơn giá 0, phải nhập tay.
ACC = [
    ("Nắp chụp 2 mảnh DN15", "Nắp chụp 2 mảnh DN15"),
    ("Nắp chụp 2 mảnh DN20", "Nắp chụp 2 mảnh DN20"),
    ("Rọ bảo vệ đầu phun DN20", "Rọ bảo vệ đầu phun DN20"),
    ("Rọ bảo vệ dầu phun DN20", "Rọ bảo vệ đầu phun DN20"),
]

def gs_record(boq, ma, ten, dvt, xx):
    return {"boq": boq, "ma": ma, "ten": ten, "xuat_xu": xx, "dvt": dvt, "don_gia": 0.0, "nguon": NGUON}

def acc_record(boq, ten):
    return {"boq": boq, "ma": "", "ten": ten, "xuat_xu": "VN/China", "dvt": "Cái", "don_gia": 0.0,
            "nguon": NGUON + " (chưa có mã, cần nhập)"}

# 1) Sửa mã đầu hướng xuống / hướng lên 1/2" bị đảo ở seed
FIX = {
    'hướng xuống 1/2': ("T-ZSTX 80-68°C Q5", "Đầu phun chữa cháy xuống lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C"),
    'hướng lên 1/2': ("T-ZSTZ 80-68°C Q5", "Đầu phun chữa cháy hướng lên phản ứng tiêu chuẩn, K=5.6, DN15, 68°C"),
}

def fix_swapped(rec):
    b = norm(rec.get("boq", ""))
    for pat, (ma, ten) in FIX.items():
        if pat in b and rec.get("ma") != ma:
            rec["ma"], rec["ten"] = ma, ten
            return True
    return False

seed = json.load(open(os.path.join(SRC_DIR, "seed_mappings.json"), encoding="utf-8"))
learned = json.load(open(os.path.join(SRC_DIR, "learned_knowledge.json"), encoding="utf-8"))

fixed = sum(fix_swapped(r) for r in seed)
new_entries = [gs_record(*g) for g in GS] + [acc_record(b, t) for b, t in ACC]
index = {norm(r["boq"]): i for i, r in enumerate(seed)}
added = updated = 0
for rec in new_entries:
    k = norm(rec["boq"])
    if k in index:
        seed[index[k]] = rec; updated += 1
    else:
        seed.append(rec); index[k] = len(seed) - 1; added += 1

# Đồng bộ learned_knowledge: app ưu tiên mục đã học, nên ghi đè các khóa vừa sửa
from datetime import datetime
now = datetime.now().isoformat(timespec="seconds")
maps = learned["mappings"]
lfixed = 0
for k, v in list(maps.items()):
    for pat, (ma, ten) in FIX.items():
        if pat in k and v.get("ma") != ma:
            v["ma"], v["ten"] = ma, ten; v["updated"] = now; lfixed += 1
for rec in new_entries:
    k = norm(rec["boq"])
    maps[k] = {"ma": rec["ma"], "ten": rec["ten"], "xuat_xu": rec["xuat_xu"], "dvt": rec["dvt"],
               "don_gia": rec["don_gia"], "count": 1, "updated": now, "nguon": rec["nguon"]}

# Kiểm tra
keys = [norm(r["boq"]) for r in seed]
assert len(keys) == len(set(keys)), "trùng khóa seed"
for r in seed:
    assert r["boq"] and r.get("ten"), "thiếu trường"

os.makedirs(OUT_DIR, exist_ok=True)
json.dump(seed, open(os.path.join(OUT_DIR, "seed_mappings.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
json.dump(learned, open(os.path.join(OUT_DIR, "learned_knowledge.json"), "w", encoding="utf-8"), ensure_ascii=False)
print(f"seed: sua_dao={fixed} them={added} cap_nhat={updated} tong={len(seed)}")
print(f"learned: sua_dao={lfixed} ghi_de_khoa_moi={len(new_entries)}")
