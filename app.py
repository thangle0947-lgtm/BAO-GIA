import streamlit as st
import pandas as pd
from fuzzywuzzy import process
import openpyxl
from openpyxl.cell.cell import MergedCell
from copy import copy
from io import BytesIO
import re
import os

# ==========================================
# CẤU HÌNH TRANG & GIAO DIỆN THƯƠNG HIỆU QN TECH
# ==========================================
st.set_page_config(
    page_title="Phần mềm Báo Giá Tự Động - QN Tech",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main {
        background-color: #f4f6f9;
    }
    .brand-header {
        background: linear-gradient(135deg, #0d233a 0%, #1e3a8a 100%);
        padding: 22px 30px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(13, 35, 58, 0.2);
        border-left: 6px solid #ff6600;
    }
    .brand-header h1 {
        margin: 0;
        font-size: 24px;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    .brand-header p {
        margin: 6px 0 0 0;
        font-size: 13px;
        color: #cbd5e1;
    }
    [data-testid="stSidebar"] {
        background-color: #0d233a;
        color: #ffffff;
    }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] span {
        color: #f1f5f9 !important;
        font-weight: 500;
    }
    [data-testid="stSidebar"] hr {
        border-color: #1e3a8a;
    }
    .stButton>button {
        background-color: #ff6600;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        border: none;
        padding: 10px 20px;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #e65c00;
        box-shadow: 0 4px 12px rgba(255, 102, 0, 0.3);
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="brand-header">
        <h1>CÔNG TY TNHH KỸ THUẬT QN TECH</h1>
        <p>Hệ thống Phân tích BOQ & Lập Báo Giá Tự Động Thông Minh (PCCC & Cơ Điện)</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# KHAI BÁO TỪ ĐIỂN CHUẨN HÓA (M&E / PCCC)
# ==========================================
BOQ_DICTIONARY = {
    "đầu phun": ["béc phun", "béc chữa cháy", "vòi phun", "đầu cứu hỏa", "đầu béc", "đầu túp", "vòi tưới", "sprinkler head", "fire sprinkler", "sprinkler", "spk"],
    "hướng xuống": ["quay xuống", "chúi xuống", "hướng đất", "chúc xuống", "pendent", "pendant", "pend"],
    "hướng lên": ["quay lên", "ngửa lên", "đứng", "quay trần", "upright", "up"],
    "hướng ngang": ["gắn vách", "vách tường", "ốp tường", "quay ngang", "phun ngang", "horizontal sidewall", "sidewall", "hsw", "sw"],
    "âm trần": ["che khuất", "giấu tăm", "nắp chụp", "nắp che", "giấu trần", "âm", "concealed", "conceal"],
    "phản ứng tiêu chuẩn": ["phản ứng chuẩn", "bầu thủy tinh chuẩn", "standard response", "sr"],
    "phản ứng nhanh": ["bầu thủy tinh nhỏ", "quick response", "fast response", "qr"],
    "van cổng": ["van chặn", "van khóa", "cổng khóa", "van đóng cắt", "gate valve", "isolation valve", "shut-off valve", "gv"],
    "ty nổi": ["tay quay ngoài", "ty ngoài", "có cần nổi", "cổ cao", "os&y", "outside stem and yoke", "rising stem"],
    "ty chìm": ["tay quay trong", "ty trong", "van nắp chụp", "chôn ngầm", "nrs", "non-rising stem"],
    "kèm công tắc giám sát": ["kèm công tắc", "có tín hiệu điện", "giám sát điện", "kèm sw", "tamper switch", "supervisory switch", "tín hiệu", "giám sát"],
    "van bướm": ["van bớm", "butterfly valve", "bfv"],
    "tay quay": ["vô lăng", "hộp số", "gear operated", "gear type", "worm gear"],
    "tay gạt": ["mỏ vịt", "tay kẹp", "lever operated", "lever type", "handle"],
    "van bi": ["van bi tay gạt", "van tay gạt", "ball valve", "van bi"],
    "van 1 chiều": ["van một chiều", "van chống chảy ngược", "van 1c", "check valve", "non-return valve", "nrv", "cv"],
    "lá lật": ["cánh lật", "bản lề", "swing check valve", "swing"],
    "cánh bướm": ["kiểu bướm", "lò xo cánh bướm", "đối trọng", "kẹp đĩa", "dual plate", "wafer check valve"],
    "lò xo": ["kiểu cối", "đĩa lò xo", "spring check valve", "lift check valve", "silent check valve"],
    "van báo động": ["van báo cháy", "cụm van báo động ướt", "trạm van báo động", "alarm check valve", "wet alarm valve", "acv", "van báo động"],
    "van xả tràn": ["van tràn ngập", "van tràn", "màng ngăn nước", "deluge valve", "deluge system", "dv"],
    "van giảm áp": ["van điều áp", "van hạ áp", "pressure reducing valve", "prv"],
    "van an toàn": ["van xả áp", "chống quá áp", "pressure relief valve", "safety valve", "relief valve"],
    "lọc y": ["y lọc", "bộ lọc chữ y", "lưới lọc bụi", "y-strainer", "strainer"],
    "rọ hút": ["luppe", "lúp bê", "rọ bơm", "foot valve", "van hút"],
    "van xả khí": ["van xả gió", "van đuổi khí", "air release valve", "air vent", "xả khí"],
    "công tắc dòng chảy": ["rơ le dòng chảy", "cảm biến dòng chảy", "báo dòng chảy", "water flow switch", "flow indicator", "flow switch"],
    "công tắc áp lực": ["rơ le áp suất", "cảm biến áp suất", "pressure switch"],
    "khớp nối mềm": ["khớp giảm chấn", "ống nối mềm", "chống rung", "cao su giảm chấn", "flexible joint", "expansion joint", "rubber joint"]
}

def chuan_hoa_boq(text):
    text_str = str(text).lower()
    for tu_chuan, danh_sach_tu_phu in BOQ_DICTIONARY.items():
        danh_sach_tu_phu = sorted(danh_sach_tu_phu, key=len, reverse=True)
        for tu_phu in danh_sach_tu_phu:
            text_str = re.sub(rf'\b{re.escape(tu_phu)}\b', tu_chuan, text_str)
    return text_str.strip()

# ==========================================
# KHAI BÁO ĐƯỜNG DẪN FILE HỆ THỐNG
# ==========================================
DATA_FILE_PATH = "BangGia.xlsx"
TEMPLATE_FILE_PATH = "MauBaoGia.xlsx"

file_data = DATA_FILE_PATH if os.path.exists(DATA_FILE_PATH) else None
file_template = TEMPLATE_FILE_PATH if os.path.exists(TEMPLATE_FILE_PATH) else None

if file_data and file_template:
    st.sidebar.success("✅ Đã kết nối Bảng giá & Mẫu báo giá.")
else:
    st.sidebar.error("❌ Thiếu file 'BangGia.xlsx' hoặc 'MauBaoGia.xlsx'!")

# ==========================================
# GIAO DIỆN BÊN TRÁI (SIDEBAR) - NHẬP LIỆU
# ==========================================
st.sidebar.markdown("### 📂 1. Tải tài liệu & Chọn Hãng")
file_boq = st.sidebar.file_uploader("Tải file BOQ Khách hàng (Excel)", type=["xlsx"])

selected_brand_sheet = None
if file_data is not None:
    try:
        xls_data = pd.ExcelFile(file_data)
        valid_sheets = [s for s in xls_data.sheet_names if not pd.read_excel(file_data, sheet_name=s).empty]
        if valid_sheets:
            selected_brand_sheet = st.sidebar.selectbox("Chọn Hãng / Sheet Bảng Giá", valid_sheets)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file Bảng giá: {e}")

selected_boq_sheet = None
if file_boq is not None:
    try:
        xls_boq = pd.ExcelFile(file_boq)
        sheet_names = xls_boq.sheet_names
        default_sheet_idx = next((i for i, name in enumerate(sheet_names) if 'thiết bị' in name.lower() or 'boq' in name.lower()), 0)
        selected_boq_sheet = st.sidebar.selectbox("Chọn Sheet chứa BOQ", sheet_names, index=default_sheet_idx)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file BOQ: {e}")

st.sidebar.markdown("### 🏢 2. Thông tin Khách hàng")
to_company = st.sidebar.text_input("To: Tên Công ty / Khách hàng", value="CÔNG TY CỔ PHẦN ĐẦU TƯ PHÁT TRIỂN MINH GLOBAL")
add_company = st.sidebar.text_input("Add: Địa chỉ khách hàng", value="415/17 Trường Chinh, Tân Bình, Hồ Chí Minh")
att_contact = st.sidebar.text_input("Att: Người liên hệ", value="Mr Kiên")
mobile_contact = st.sidebar.text_input("Mobile: Số điện thoại", value="0909123456")
email_contact = st.sidebar.text_input("Email", value="kien@minhglobal.vn")
ref_content = st.sidebar.text_input("Ref: Nội dung báo giá", value="BÁO GIÁ THIẾT BỊ HỆ THỐNG PCCC")

st.sidebar.markdown("### 📋 3. Quản lý Báo giá & Dự án")
no_quote = st.sidebar.text_input("No: Số báo giá", value="Q2026/2108/PCCC/MG")
date_quote = st.sidebar.text_input("Date: Ngày báo giá", value="01/09/2026")
validity_quote = st.sidebar.text_input("Validity: Hiệu lực", value="30/09/2026")
project_name = st.sidebar.text_input("Project: Tên Dự án", value="NHÀ MÁY TRẦN ĐỨC / BỆNH VIỆN QUÂN Y 7B")
output_filename = st.sidebar.text_input("Tên file xuất ra:", value="BaoGia_QNTech_HoanChinh.xlsx")

if 'df_editable' not in st.session_state:
    st.session_state.df_editable = None

# ==========================================
# CHUẨN BỊ DỮ LIỆU THAM KHẢO BẢNG GIÁ
# ==========================================
list_options = ["-- Chọn mã từ bảng giá --", "-- Dùng tên gốc BOQ (Không có mã) --"]
dict_mapping_price = {}

if file_data and selected_brand_sheet:
    try:
        df_ref_full = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
        df_ref_full.columns = [str(c).strip() for c in df_ref_full.columns]
        
        ref_code_col = next((c for c in df_ref_full.columns if c.lower() in ['mã sp', 'mã sản phẩm', 'code', 'mã vật tư']), 'Mã SP')
        ref_name_col = next((c for c in df_ref_full.columns if c.lower() in ['tên chuẩn', 'tên thiết bị', 'tên hàng', 'tên vật tư']), 'Tên chuẩn')
        ref_price_col = next((c for c in df_ref_full.columns if c.lower() in ['đơn giá', 'price', 'giá']), 'Đơn giá')
        ref_unit_col = next((c for c in df_ref_full.columns if c.lower() in ['đvt', 'đơn vị tính', 'unit']), 'ĐVT')
        ref_origin_col = next((c for c in df_ref_full.columns if c.lower() in ['xuất xứ', 'origin']), 'XUẤT XỨ')

        df_ref_full = df_ref_full.dropna(subset=[ref_code_col])
        df_ref_full[ref_code_col] = df_ref_full[ref_code_col].astype(str).str.strip()
        df_ref_full = df_ref_full[~df_ref_full[ref_code_col].isin(['nan', 'None', '', '-'])]
        df_ref_full[ref_name_col] = df_ref_full[ref_name_col].astype(str).str.strip()

        df_ref_full['Option_Display'] = df_ref_full[ref_code_col] + " - " + df_ref_full[ref_name_col]
        df_ref_full = df_ref_full.drop_duplicates(subset=['Option_Display'])
        
        list_options = ["-- Chọn mã từ bảng giá --", "-- Dùng tên gốc BOQ (Không có mã) --"] + df_ref_full['Option_Display'].tolist()
        
        def safe_float_convert(val):
            try:
                if val is None or pd.isna(val):
                    return 0.0
                s = str(val).replace(',', '').strip()
                if not s or s == '-' or s.lower() == 'nan':
                    return 0.0
                return float(s)
            except ValueError:
                return 0.0

        for _, r in df_ref_full.iterrows():
            code_val = str(r[ref_code_col])
            price_val = safe_float_convert(r.get(ref_price_col, 0))
            dict_mapping_price[code_val] = {
                'Tên chuẩn': str(r.get(ref_name_col, '')),
                'Đơn giá': price_val,
                'ĐVT': str(r.get(ref_unit_col, 'Cái')),
                'XUẤT XỨ': str(r.get(ref_origin_col, 'LS/ VN'))
            }
    except Exception as e:
        st.sidebar.error(f"Lỗi tải danh mục bảng giá: {e}")

# ==========================================
# XỬ LÝ DỮ LIỆU BAN ĐẦU
# ==========================================
st.markdown("### ⚡ Bảng Điều Khiển Xử Lý BOQ")
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    run_process = st.button("🚀 Phân Tích BOQ", type="primary")

if run_process:
    if not (file_data and file_boq and file_template and selected_brand_sheet):
        st.error("⚠️ Vui lòng tải lên file BOQ và kiểm tra đầy đủ Bảng Giá / Mẫu Báo Giá trong Sidebar!")
    else:
        try:
            with st.spinner("Đang đọc dữ liệu, bóc tách thông số K, Nhiệt độ và phân tích BOQ..."):
                
                # --- 1. ĐỌC DỮ LIỆU BẢNG GIÁ ---
                df_data = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
                df_data.columns = [str(c).strip() for c in df_data.columns]
                
                data_code_col = next((c for c in df_data.columns if c.lower() in ['mã sp', 'mã sản phẩm', 'code', 'mã vật tư']), 'Mã SP')
                data_name_col = next((c for c in df_data.columns if c.lower() in ['tên chuẩn', 'tên thiết bị', 'tên hàng', 'tên vật tư', 'diễn giải']), 'Tên chuẩn')
                
                if data_name_col != 'Tên chuẩn':
                    df_data = df_data.rename(columns={data_name_col: 'Tên chuẩn'})
                
                df_data = df_data.dropna(subset=[data_code_col])
                df_data[data_code_col] = df_data[data_code_col].astype(str).str.strip()
                df_data = df_data[~df_data[data_code_col].isin(['nan', 'None', '', '-'])]
                
                df_data['Tên chuẩn'] = df_data['Tên chuẩn'].astype(str).str.strip()
                danh_sach_chuan = df_data['Tên chuẩn'].tolist()

                # --- 2. ĐỌC DỮ LIỆU BOQ ---
                df_boq = pd.read_excel(file_boq, sheet_name=selected_boq_sheet)
                df_boq.columns = [str(c).strip() for c in df_boq.columns]
                
                name_col_found = next((c for c in df_boq.columns if c.lower() in ['tên chuẩn', 'tên thiết bị boq', 'diễn giải', 'tên hàng', 'tên vật tư', 'tên thiết bị']), df_boq.columns[1])
                qty_col_found = next((c for c in df_boq.columns if c.lower() in ['khối lượng', 'số lượng', 'sl']), 'Số lượng')
                k_col_boq = next((c for c in df_boq.columns if c.lower() in ['k', 'hệ số k']), None)
                temp_col_boq = next((c for c in df_boq.columns if c.lower() in ['nhiệt độ', 'temp', 'nhiet do']), None)

                df_boq = df_boq.rename(columns={
                    name_col_found: 'Tên thiết bị BOQ',
                    qty_col_found: 'Số lượng'
                })

                df_boq['K_BOQ'] = df_boq[k_col_boq] if k_col_boq else 8.0
                df_boq['Temp_BOQ'] = df_boq[temp_col_boq] if temp_col_boq else '68°C'

                df_boq = df_boq.dropna(subset=['Tên thiết bị BOQ'])
                df_boq['Tên thiết bị BOQ'] = df_boq['Tên thiết bị BOQ'].astype(str).str.replace('\u3000', ' ').str.strip()
                df_boq['Số lượng'] = pd.to_numeric(df_boq['Số lượng'], errors='coerce').fillna(0)

                def clean_k_value(val):
                    match = re.search(r'([0-9]+(?:\.[0-9]*)?)', str(val))
                    return float(match.group(1)) if match else 8.0

                def clean_temp_value(val):
                    match = re.search(r'(\d+)', str(val))
                    return int(match.group(1)) if match else 68

                df_boq['K_Val'] = df_boq['K_BOQ'].apply(clean_k_value)
                df_boq['Temp_Val'] = df_boq['Temp_BOQ'].apply(clean_temp_value)

                def extract_dn(text):
                    match = re.search(r'DN\s*[:=]?\s*(\d+)', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 15

                def extract_k_factor(text):
                    match = re.search(r'K\s*[:=]?\s*([0-9]+(?:\.[0-9]*)?)', str(text), re.IGNORECASE)
                    return float(match.group(1)) if match else 5.5

                def extract_temperature(text):
                    match = re.search(r't\s*[:=]?\s*(\d+)', str(text), re.IGNORECASE)
                    if not match:
                        match = re.search(r'(\d+)\s*°\s*c', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 68

                def map_ten_va_quy_tac_chuan(row):
                    text_chuan = chuan_hoa_boq(row['Tên thiết bị BOQ'])
                    k_val = row['K_Val']
                    temp_val = row['Temp_Val']
                    dn_val = extract_dn(row['Tên thiết bị BOQ'])
                    
                    if 'đầu phun' in text_chuan:
                        is_nhanh = 'nhanh' in text_chuan
                        response_str = "phản ứng nhanh" if is_nhanh else "phản ứng tiêu chuẩn"
                        if 'hướng lên' in text_chuan:
                            huong_str = "hướng lên"
                        elif 'hướng ngang' in text_chuan:
                            huong_str = "hướng ngang"
                        else:
                            huong_str = "hướng xuống"
                            
                        if dn_val == 15:
                            if k_val > 11.2: dn_val = 25
                            elif k_val >= 8.0: dn_val = 20
                            
                        return f"Đầu phun chữa cháy hướng {huong_str} {response_str}, K={k_val:g}, DN{dn_val}, {temp_val}°C"
                    return text_chuan

                df_boq['Tên BOQ Chuẩn Hóa'] = df_boq.apply(map_ten_va_quy_tac_chuan, axis=1)

                def tim_ma_phu_hop_chinh_xac(row):
                    ten_chuan_boq = str(row['Tên BOQ Chuẩn Hóa']).lower()
                    k_target = float(row['K_Val'])
                    temp_target = int(row['Temp_Val'])
                    dn_val = extract_dn(row['Tên thiết bị BOQ'])
                    
                    if dn_val == 15:
                        if k_target > 11.2: dn_val = 25
                        elif k_target >= 8.0: dn_val = 20

                    if 'van báo động' in ten_chuan_boq:
                        danh_sach_loc = [item for item in danh_sach_chuan if 'van báo động' in str(item).lower() or 'alarm check valve' in str(item).lower()]
                        if danh_sach_loc:
                            if dn_val > 0:
                                for item in danh_sach_loc:
                                    if f"dn{dn_val}" in str(item).lower() or f"dn {dn_val}" in str(item).lower():
                                        return item
                            return danh_sach_loc[0]
                        return "Cần check thủ công"

                    if 'van bi' in ten_chuan_boq:
                        danh_sach_loc = [item for item in danh_sach_chuan if 'van bi' in str(item).lower() or 'ball valve' in str(item).lower()]
                        if danh_sach_loc:
                            if dn_val > 0:
                                for item in danh_sach_loc:
                                    if f"dn{dn_val}" in str(item).lower() or f"dn {dn_val}" in str(item).lower():
                                        return item
                            return danh_sach_loc[0]
                        return "Cần check thủ công"

                    if 'rọ hút' in ten_chuan_boq:
                        danh_sach_loc = [item for item in danh_sach_chuan if 'rọ hút' in str(item).lower() or 'luppe' in str(item).lower() or 'foot valve' in str(item).lower() or 'rọ bơm' in str(item).lower()]
                        if danh_sach_loc:
                            if dn_val > 0:
                                for item in danh_sach_loc:
                                    if f"dn{dn_val}" in str(item).lower() or f"dn {dn_val}" in str(item).lower():
                                        return item
                            return danh_sach_loc[0]
                        return "Cần check thủ công"

                    if 'đầu phun' in ten_chuan_boq:
                        is_nhanh_boq = 'nhanh' in ten_chuan_boq
                        is_len_boq = 'hướng lên' in ten_chuan_boq
                        is_ngang_boq = 'hướng ngang' in ten_chuan_boq
                        is_xuong_boq = 'hướng xuống' in ten_chuan_boq or (not is_len_boq and not is_ngang_boq)

                        danh_sach_sp_phu_hop = []
                        for item in danh_sach_chuan:
                            item_str = str(item).lower()
                            if 'đầu phun' in item_str or 'sprinkler' in item_str:
                                if is_nhanh_boq != ('nhanh' in item_str):
                                    continue
                                item_k = extract_k_factor(item_str)
                                item_temp = extract_temperature(item_str)
                                item_dn = extract_dn(item_str)
                                
                                if abs(k_target - item_k) >= 0.1 or temp_target != item_temp:
                                    continue
                                if dn_val > 0 and dn_val != item_dn:
                                    continue

                                is_len_item = 'hướng lên' in item_str or 'upright' in item_str
                                is_ngang_item = 'hướng ngang' in item_str or 'sidewall' in item_str
                                is_xuong_item = 'hướng xuống' in item_str or 'pendent' in item_str or (not is_len_item and not is_ngang_item)

                                if is_len_boq == is_len_item and is_ngang_boq == is_ngang_item and is_xuong_boq == is_xuong_item:
                                    danh_sach_sp_phu_hop.append(item)
                        if danh_sach_sp_phu_hop:
                            return danh_sach_sp_phu_hop[0]
                        return "Cần check thủ công"

                    if 'van cổng' in ten_chuan_boq:
                        is_osy = 'ty nổi' in ten_chuan_boq or 'os&y' in ten_chuan_boq
                        has_switch = 'tín hiệu' in ten_chuan_boq or 'giám sát' in ten_chuan_boq or 'công tắc' in ten_chuan_boq or 'sw' in ten_chuan_boq
                        
                        danh_sach_loc = []
                        for item in danh_sach_chuan:
                            item_lower = str(item).lower()
                            if 'van cổng' not in item_lower and 'cổng' not in item_lower:
                                continue
                            if is_osy != ('os&y' in item_lower or 'ty nổi' in item_lower):
                                continue
                            if has_switch != ('giám sát' in item_lower or 'tín hiệu' in item_lower or 'công tắc' in item_lower or 'tamper' in item_lower or 'switch' in item_lower):
                                continue
                            danh_sach_loc.append(item)
                        
                        if danh_sach_loc:
                            if dn_val > 0:
                                for item in danh_sach_loc:
                                    if f"dn{dn_val}" in str(item).lower() or f"dn {dn_val}" in str(item).lower():
                                        return item
                            return danh_sach_loc[0]
                        return "Cần check thủ công"

                    match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                    return match if score > 40 else "Cần check thủ công"

                df_boq['Tên Map'] = df_boq.apply(tim_ma_phu_hop_chinh_xac, axis=1)
                df_result = pd.merge(df_boq, df_data, left_on='Tên Map', right_on='Tên chuẩn', how='left')
                
                default_selections = []
                for _, r in df_result.iterrows():
                    m = str(r.get('Mã SP', ''))
                    match_opt = next((opt for opt in list_options if opt.startswith(m + " -")), "-- Chọn mã từ bảng giá --")
                    default_selections.append(match_opt)
                
                df_result['Chọn / Đổi Mã SP'] = default_selections
                st.session_state.df_editable = df_result
                st.success("🎉 Đã Phân Tích Thành Công!")

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")

# ==========================================
# GIAO DIỆN HIỂN THỊ & CHỈNH SỬA RÚT GỌN
# ==========================================
if st.session_state.df_editable is not None:
    df_current = st.session_state.df_editable
    
    total_items = len(df_current)
    manual_count = len(df_current[(df_current['Tên Map'] == 'Cần check thủ công') | (df_current['Mã SP'].isna()) | (df_current['Mã SP'] == '') | (df_current['Mã SP'] == 'None')])
    auto_count = total_items - manual_count

    m1, m2, m3 = st.columns(3)
    m1.metric("Tổng số thiết bị BOQ", f"{total_items} mục")
    m2.metric("Nhận diện tự động chuẩn", f"{auto_count} mục", delta="Thành công")
    m3.metric("Cần kiểm tra thủ công", f"{manual_count} mục", delta=f"-{manual_count}" if manual_count > 0 else "0", delta_color="inverse")

    st.markdown("### 📋 Bảng Tra Cứu & Chỉnh Sửa Rút Gọn")
    st.info("💡 **Hướng dẫn:** Bảng hiển thị rút gọn các cột chính. Bạn có thể chọn lại mã hàng, hoặc chọn **'-- Dùng tên gốc BOQ (Không có mã) --'** để hệ thống tự điền tên gốc từ BOQ vào báo giá.")

    if 'Chọn / Đổi Mã SP' not in df_current.columns:
        df_current['Chọn / Đổi Mã SP'] = "-- Chọn mã từ bảng giá --"

    edited_df = st.data_editor(
        df_current,
        column_config={
            "Chọn / Đổi Mã SP": st.column_config.SelectboxColumn(
                "Chọn / Đổi Mã SP",
                options=list_options,
                required=True,
            ),
            "Tên thiết bị BOQ": st.column_config.TextColumn("Tên thiết bị BOQ", disabled=True),
            "Tên Map": st.column_config.TextColumn("Tên Map", disabled=True),
            "Mã SP": st.column_config.TextColumn("Mã SP"),
            "Số lượng": st.column_config.NumberColumn("Số lượng", min_value=0),
            "Đơn giá": st.column_config.NumberColumn("Đơn giá (VNĐ)", min_value=0, format="%d"),
        },
        column_order=["Chọn / Đổi Mã SP", "Tên thiết bị BOQ", "Tên Map", "Mã SP", "Số lượng", "Đơn giá"],
        hide_index=False,
        key="master_grid_editor_compact"
    )

    # Xử lý cập nhật thông tin khi thay đổi lựa chọn trong dropdown
    for idx, row in edited_df.iterrows():
        selected_option = row.get('Chọn / Đổi Mã SP')
        if selected_option == "-- Dùng tên gốc BOQ (Không có mã) --":
            edited_df.at[idx, 'Mã SP'] = ""
            edited_df.at[idx, 'Tên chuẩn'] = row.get('Tên thiết bị BOQ', '')
            edited_df.at[idx, 'Đơn giá'] = 0.0
            edited_df.at[idx, 'ĐVT'] = "Cái"
            edited_df.at[idx, 'XUẤT XỨ'] = ""
        elif selected_option and selected_option != "-- Chọn mã từ bảng giá --":
            chosen_code = selected_option.split(" - ")[0].strip()
            if chosen_code in dict_mapping_price and chosen_code != str(row.get('Mã SP', '')):
                info = dict_mapping_price[chosen_code]
                edited_df.at[idx, 'Mã SP'] = chosen_code
                edited_df.at[idx, 'Tên chuẩn'] = info['Tên chuẩn']
                edited_df.at[idx, 'Đơn giá'] = info['Đơn giá']
                edited_df.at[idx, 'ĐVT'] = info['ĐVT']
                edited_df.at[idx, 'XUẤT XỨ'] = info['XUẤT XỨ']

    st.session_state.df_editable = edited_df

    st.markdown("---")
    if st.button("📥 Xuất File Báo Giá Hoàn Chỉnh", type="primary"):
        try:
            with st.spinner("Đang tổng hợp dữ liệu và điền vào mẫu báo giá QN Tech..."):
                
                df_final = st.session_state.df_editable.copy()
                df_final['Thành tiền'] = pd.to_numeric(df_final['Số lượng'], errors='coerce').fillna(0) * pd.to_numeric(df_final['Đơn giá'], errors='coerce').fillna(0)

                wb = openpyxl.load_workbook(TEMPLATE_FILE_PATH)
                ws = wb.active

                def safe_write(row, col, val):
                    cell = ws.cell(row=row, column=col)
                    if not isinstance(cell, MergedCell):
                        cell.value = val

                safe_write(6, 2, to_company)        
                safe_write(7, 2, add_company)       
                safe_write(8, 2, att_contact)       
                safe_write(9, 2, mobile_contact)    
                safe_write(10, 2, email_contact)    
                safe_write(11, 2, ref_content)      

                safe_write(6, 7, no_quote)          
                safe_write(7, 7, date_quote)        
                safe_write(9, 7, validity_quote)    
                safe_write(10, 7, project_name)   

                start_row = 14  
                num_items = len(df_final)

                footer_row = start_row + 1
                for r in range(start_row, start_row + 1000):
                    v2 = str(ws.cell(row=r, column=2).value or "").strip().lower()
                    v3 = str(ws.cell(row=r, column=3).value or "").strip().lower()
                    if 'cộng' in v2 or 'cộng' in v3 or 'total' in v2:
                        footer_row = r
                        break

                available_rows = footer_row - start_row

                if num_items <= available_rows:
                    for r in range(start_row + num_items, footer_row):
                        ws.row_dimensions[r].hidden = True
                else:
                    rows_to_add = num_items - available_rows
                    ws.insert_rows(footer_row, amount=rows_to_add)
                    src_row = footer_row - 1
                    for r_offset in range(rows_to_add):
                        target_r = footer_row + r_offset
                        for c in range(1, 12):
                            src_cell = ws.cell(row=src_row, column=c)
                            tgt_cell = ws.cell(row=target_r, column=c)
                            if src_cell.has_style:
                                tgt_cell.font = copy(src_cell.font)
                                tgt_cell.border = copy(src_cell.border)
                                tgt_cell.fill = copy(src_cell.fill)
                                tgt_cell.number_format = src_cell.number_format
                                tgt_cell.alignment = copy(src_cell.alignment)

                for index, row_data in df_final.iterrows():
                    current_row = start_row + index
                    safe_write(current_row, 1, index + 1)                
                    safe_write(current_row, 2, row_data.get('Mã SP', ''))        
                    safe_write(current_row, 3, row_data.get('Tên chuẩn', ''))    
                    safe_write(current_row, 4, row_data.get('XUẤT XỨ', ''))      
                    safe_write(current_row, 5, row_data.get('ĐVT', 'Cái') if pd.notna(row_data.get('ĐVT')) else 'Cái')
                    safe_write(current_row, 6, row_data.get('Số lượng', 0))      
                    safe_write(current_row, 7, row_data.get('Đơn giá', 0))        
                    safe_write(current_row, 8, row_data.get('Thành tiền', 0))    
                    safe_write(current_row, 9, '')                               

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                final_name = output_filename if output_filename.endswith('.xlsx') else output_filename + '.xlsx'

                st.download_button(
                    label="💾 TẢI XUỐNG FILE BÁO GIÁ EXCEL",
                    data=output,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo file Excel: {e}")
