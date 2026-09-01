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

# Custom CSS đồng bộ nhận diện logo QN Tech (Xanh đậm #0D233A & Cam #FF6600)
st.markdown("""
    <style>
    .main {
        background-color: #f4f6f9;
    }
    
    /* Header Thương hiệu */
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

    /* Sidebar Styling */
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

    /* Nút bấm chính */
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
    
    /* Tiêu đề mục */
    h3 {
        color: #0d233a;
        font-weight: 700;
        border-bottom: 2px solid #ff6600;
        padding-bottom: 6px;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# Banner Giao diện
st.markdown("""
    <div class="brand-header">
        <h1>CÔNG TY TNHH KỸ THUẬT QN TECH</h1>
        <p>Hệ thống Phân tích BOQ & Lập Báo Giá Tự Động Thông Minh (PCCC & Cơ Điện)</p>
    </div>
""", unsafe_allow_html=True)

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

# Khởi tạo Session State
if 'df_raw_result' not in st.session_state:
    st.session_state.df_raw_result = None
if 'df_manual_edited' not in st.session_state:
    st.session_state.df_manual_edited = None

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
            with st.spinner("Đang phân tích BOQ theo hệ thống 3 tầng thông minh..."):
                
                # --- ĐỌC DỮ LIỆU BẢNG GIÁ ---
                df_data = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
                df_data.columns = [str(c).strip() for c in df_data.columns]
                
                data_name_col = next((c for c in df_data.columns if c.lower() in ['tên chuẩn', 'tên thiết bị', 'tên hàng', 'tên vật tư', 'diễn giải']), None)
                if data_name_col != 'Tên chuẩn':
                    df_data = df_data.rename(columns={data_name_col: 'Tên chuẩn'})
                df_data['Tên chuẩn'] = df_data['Tên chuẩn'].astype(str).str.strip()
                df_data = df_data.dropna(subset=['Tên chuẩn'])
                danh_sach_chuan = df_data['Tên chuẩn'].tolist()

                # --- ĐỌC DỮ LIỆU BOQ ---
                df_boq = pd.read_excel(file_boq, sheet_name=selected_boq_sheet)
                df_boq.columns = [str(c).strip() for c in df_boq.columns]
                
                name_col_found = next((c for c in df_boq.columns if c.lower() in ['tên thiết bị boq', 'diễn giải', 'tên hàng', 'tên vật tư', 'tên thiết bị', 'tên công tác']), None)
                if not name_col_found and len(df_boq.columns) > 1:
                    name_col_found = df_boq.columns[1]
                
                if name_col_found != 'Tên thiết bị BOQ':
                    df_boq = df_boq.rename(columns={name_col_found: 'Tên thiết bị BOQ'})

                qty_col_found = next((c for c in df_boq.columns if c.lower() in ['khối lượng', 'số lượng', 'sl']), None)
                if qty_col_found:
                    df_boq = df_boq.rename(columns={qty_col_found: 'Số lượng'})
                else:
                    df_boq['Số lượng'] = 0

                df_boq = df_boq.dropna(subset=['Tên thiết bị BOQ'])
                df_boq['Tên thiết bị BOQ'] = df_boq['Tên thiết bị BOQ'].astype(str).str.replace('\u3000', ' ').str.strip()
                df_boq['Số lượng'] = pd.to_numeric(df_boq['Số lượng'], errors='coerce').fillna(0)
                df_boq = df_boq.groupby('Tên thiết bị BOQ', sort=False, as_index=False)['Số lượng'].sum()

                # --- HÀM HỖ TRỢ TRÍCH XUẤT THÔNG SỐ (REGEX XỬ LÝ DẤU =) ---
                def extract_dn(text):
                    match = re.search(r'DN\s*[:=]?\s*(\d+)', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 0

                def extract_k_factor(text):
                    match = re.search(r'K\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)', str(text), re.IGNORECASE)
                    return float(match.group(1)) if match else 5.5

                def extract_temperature(text):
                    match = re.search(r't\s*[:=]?\s*(\d+)', str(text), re.IGNORECASE)
                    if not match:
                        match = re.search(r'(\d+)\s*°\s*c', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 68

                def map_ten_va_quy_tac(ten_boq):
                    text_lower = str(ten_boq).lower()
                    dn_val = extract_dn(ten_boq)
                    k_val = extract_k_factor(text_lower)
                    temp_val = extract_temperature(text_lower)

                    if 'đầu phun' in text_lower or 'sprinkler' in text_lower:
                        is_nhanh = 'nhanh' in text_lower
                        response_str = "phản ứng nhanh" if is_nhanh else "phản ứng tiêu chuẩn"
                        
                        if 'hướng lên' in text_lower or 'upright' in text_lower:
                            huong_str = "hướng lên"
                        elif 'hướng ngang' in text_lower or 'sidewall' in text_lower:
                            huong_str = "hướng ngang"
                        else:
                            huong_str = "hướng xuống"
                            
                        if dn_val == 0:
                            if k_val >= 11.2: dn_val = 25
                            elif k_val >= 8.0: dn_val = 20
                            else: dn_val = 15

                        return f"Đầu phun chữa cháy hướng {huong_str} {response_str}, K={k_val:g}, DN{dn_val}, {temp_val}°C"

                    if 'silicone' in text_lower or 'chống cháy' in text_lower or 'cap' in text_lower or 'cáp' in text_lower:
                        has_khong_nhieu = 'không chống nhiễu' in text_lower or 'ko chống nhiễu' in text_lower
                        has_chống_nhieu = ('chống nhiễu' in text_lower or 'shield' in text_lower or 'st' in text_lower) and not has_khong_nhieu
                        shield_text = "chống cháy chống nhiễu" if has_chống_nhieu else "chống cháy không chống nhiễu"

                        core_match = re.search(r'(\d+)\s*(?:lõi|c\b|core|x)', text_lower)
                        num_cores = int(core_match.group(1)) if core_match else 2

                        size_val = "1.5"
                        if '2.5' in text_lower or '25' in text_lower: size_val = "2.5"
                        elif '1.5' in text_lower or '15' in text_lower: size_val = "1.5"
                        elif '1.0' in text_lower: size_val = "1.0"
                        elif '0.5' in text_lower: size_val = "0.50"
                        elif '4.0' in text_lower: size_val = "4.0"

                        return f"Cáp tín hiệu điều khiển Silicone {shield_text}, {num_cores} lõi, {size_val}mm2"

                    if 'khói' in text_lower and ('địa chỉ' in text_lower or 'đ/c' in text_lower): return "Đầu báo khói quang địa chỉ kèm đế"
                    if 'giám sát đầu báo thường' in text_lower or '(zm)' in text_lower: return "Module 2 ngõ vào giám sát đầu báo thường"
                    if 'nút nhấn' in text_lower and 'địa chỉ' in text_lower: return "Nút nhấn báo cháy địa chỉ"
                    if ('còi đèn' in text_lower or 'chuông đèn' in text_lower) and 'địa chỉ' in text_lower: return "Còi đèn chớp báo cháy địa chỉ kèm đế"
                    return ten_boq

                # THUẬT TOÁN MAPPING 3 TẦNG THÔNG MINH
                def tim_ma_phu_hop(ten_chuan_boq, ten_goc_boq):
                    dn_val = extract_dn(ten_goc_boq)
                    k_val = extract_k_factor(ten_goc_boq)
                    temp_val = extract_temperature(ten_goc_boq)
                    if dn_val == 0:
                        if k_val >= 11.2: dn_val = 25
                        elif k_val >= 8.0: dn_val = 20
                        else: dn_val = 15

                    text_lower = str(ten_goc_boq).lower()

                    if 'đầu phun' in text_lower or 'sprinkler' in text_lower:
                        is_nhanh_boq = 'nhanh' in text_lower
                        is_len_boq = 'hướng lên' in text_lower or 'upright' in text_lower
                        is_ngang_boq = 'hướng ngang' in text_lower or 'sidewall' in text_lower
                        is_xuong_boq = not is_len_boq and not is_ngang_boq

                        # TẦNG 1: Khớp tuyệt đối 100%
                        for item in danh_sach_chuan:
                            item_str = str(item).lower()
                            if 'đầu phun' in item_str or 'sprinkler' in item_str:
                                if ('nhanh' in item_str) == is_nhanh_boq:
                                    if abs(k_val - extract_k_factor(item_str)) < 0.1:
                                        if temp_val == extract_temperature(item_str):
                                            item_dn = extract_dn(item_str)
                                            if item_dn == 0:
                                                item_k = extract_k_factor(item_str)
                                                item_dn = 25 if item_k >= 11.2 else (20 if item_k >= 8.0 else 15)
                                            is_len_item = 'hướng lên' in item_str or 'upright' in item_str
                                            is_ngang_item = 'hướng ngang' in item_str or 'sidewall' in item_str
                                            is_xuong_item = not is_len_item and not is_ngang_item
                                            if (is_len_boq == is_len_item) and (is_ngang_boq == is_ngang_item) and (is_xuong_boq == is_xuong_item):
                                                if dn_val == item_dn:
                                                    return item

                        # TẦNG 2: Khớp linh hoạt
                        best_match = None
                        min_diff = 999.0
                        for item in danh_sach_chuan:
                            item_str = str(item).lower()
                            if 'đầu phun' in item_str or 'sprinkler' in item_str:
                                if ('nhanh' in item_str) == is_nhanh_boq:
                                    is_len_item = 'hướng lên' in item_str or 'upright' in item_str
                                    is_ngang_item = 'hướng ngang' in item_str or 'sidewall' in item_str
                                    is_xuong_item = not is_len_item and not is_ngang_item
                                    if (is_len_boq == is_len_item) and (is_ngang_boq == is_ngang_item) and (is_xuong_boq == is_xuong_item):
                                        item_k = extract_k_factor(item_str)
                                        diff = abs(k_val - item_k)
                                        if diff < min_diff:
                                            min_diff = diff
                                            best_match = item
                        if best_match is not None:
                            return best_match

                    # TẦNG 3: Fallback FuzzyWuzzy
                    match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                    if score > 40:
                        return match
                    
                    return "Cần check thủ công"

                df_boq['Tên BOQ Chuẩn Hóa'] = df_boq['Tên thiết bị BOQ'].apply(map_ten_va_quy_tac)
                df_boq['Tên Map'] = df_boq.apply(lambda row: tim_ma_phu_hop(row['Tên BOQ Chuẩn Hóa'], row['Tên thiết bị BOQ']), axis=1)
                
                df_result = pd.merge(df_boq, df_data, left_on='Tên Map', right_on='Tên chuẩn', how='left')
                df_result['Đơn giá'] = pd.to_numeric(df_result['Đơn giá'], errors='coerce').fillna(0)
                
                st.session_state.df_raw_result = df_result
                st.success("🎉 Phân tích BOQ thành công bằng hệ thống thông minh!")

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")

# ==========================================
# GIAO DIỆN HIỂN THỊ KẾT QUẢ & TRA CỨU THỦ CÔNG
# ==========================================
if st.session_state.df_raw_result is not None:
    df_current = st.session_state.df_raw_result
    
    total_items = len(df_current)
    manual_count = len(df_current[df_current['Tên Map'] == 'Cần check thủ công'])
    auto_count = total_items - manual_count

    # Thẻ thống kê trực quan
    m1, m2, m3 = st.columns(3)
    m1.metric("Tổng số thiết bị BOQ", f"{total_items} mục")
    m2.metric("Nhận diện tự động chuẩn", f"{auto_count} mục", delta="Thành công")
    m3.metric("Cần kiểm tra thủ công", f"{manual_count} mục", delta=f"-{manual_count}" if manual_count > 0 else "0", delta_color="inverse")

    st.subheader("🔍 Bảng Đối Chiếu Kết Quả Mapping")
    display_compare_df = df_current[['Tên thiết bị BOQ', 'Tên BOQ Chuẩn Hóa', 'Tên Map', 'Mã SP', 'Số lượng', 'Đơn giá']].copy()
    st.dataframe(display_compare_df, use_container_width=True)

    st.subheader("📋 Tra Cứu & Hiệu Chỉnh Thông Minh (Chọn Mã SP Trực Tiếp)")
    
    # Chuẩn bị dữ liệu từ bảng giá hãng để tạo popup/dropdown tra cứu
    df_ref_full = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
    df_ref_full.columns = [str(c).strip() for c in df_ref_full.columns]
    
    ref_code_col = next((c for c in df_ref_full.columns if c.lower() in ['mã sp', 'mã sản phẩm', 'code', 'mã vật tư']), 'Mã SP')
    ref_name_col = next((c for c in df_ref_full.columns if c.lower() in ['tên chuẩn', 'tên thiết bị', 'tên hàng', 'tên vật tư']), 'Tên chuẩn')
    ref_price_col = next((c for c in df_ref_full.columns if c.lower() in ['đơn giá', 'price', 'giá']), 'Đơn giá')
    ref_unit_col = next((c for c in df_ref_full.columns if c.lower() in ['đvt', 'đơn vị tính', 'unit']), 'ĐVT')
    ref_origin_col = next((c for c in df_ref_full.columns if c.lower() in ['xuất xứ', 'origin']), 'XUẤT XỨ')

    # Tạo danh sách lựa chọn cho dropdown
    df_ref_full['Option_Display'] = df_ref_full[ref_code_col].astype(str) + " - " + df_ref_full[ref_name_col].astype(str)
    list_options = ["-- Chọn mã từ bảng giá --"] + df_ref_full['Option_Display'].tolist()
    
    dict_mapping_price = {}
    for _, r in df_ref_full.iterrows():
        code_val = str(r[ref_code_col])
        dict_mapping_price[code_val] = {
            'Tên chuẩn': str(r.get(ref_name_col, '')),
            'Đơn giá': float(r.get(ref_price_col, 0)) if pd.notna(r.get(ref_price_col)) else 0.0,
            'ĐVT': str(r.get(ref_unit_col, 'Cái')),
            'XUẤT XỨ': str(r.get(ref_origin_col, 'LS/ VN'))
        }

    mask_manual = (df_current['Tên Map'] == 'Cần check thủ công') | (df_current['Mã SP'].isna()) | (df_current['Mã SP'] == '')
    df_auto = df_current[~mask_manual].copy()
    df_manual = df_current[mask_manual].copy()

    if not df_manual.empty:
        if 'Bỏ qua' not in df_manual.columns:
            df_manual.insert(0, 'Bỏ qua', False)
        if 'Chọn Mã Từ Bảng Giá' not in df_manual.columns:
            df_manual['Chọn Mã Từ Bảng Giá'] = "-- Chọn mã từ bảng giá --"
        if 'XUẤT XỨ' not in df_manual.columns:
            df_manual['XUẤT XỨ'] = 'LS/ VN'
        if 'ĐVT' not in df_manual.columns:
            df_manual['ĐVT'] = 'Cái'

        st.info("💡 **Mẹo:** Nhấp vào ô trong cột **'Chọn Mã Từ Bảng Giá'** để mở hộp tìm kiếm dạng dropdown. Bạn có thể gõ từ khóa (mã hoặc tên) để lọc nhanh, sau đó chọn mã hãng phù hợp — hệ thống sẽ tự động điền đầy đủ Mã SP, Tên chuẩn, Đơn giá và ĐVT.")

        edited_manual_df = st.data_editor(
            df_manual[['Bỏ qua', 'Tên thiết bị BOQ', 'Chọn Mã Từ Bảng Giá', 'Mã SP', 'Tên chuẩn', 'XUẤT XỨ', 'ĐVT', 'Số lượng', 'Đơn giá']],
            column_config={
                "Chọn Mã Từ Bảng Giá": st.column_config.SelectboxColumn(
                    "🔍 Tra Cứu & Chọn Mã SP Từ Bảng Giá",
                    help="Nhấp để tìm kiếm hoặc chọn mã sản phẩm trong bảng giá hãng",
                    options=list_options,
                    required=True
                )
            },
            num_rows="fixed",
            key="manual_editor_search",
            use_container_width=True
        )

        # Tự động đồng bộ thông tin khi người dùng chọn mã từ dropdown
        for idx, row in edited_manual_df.iterrows():
            selected_option = row['Chọn Mã Từ Bảng Giá']
            if selected_option and selected_option != "-- Chọn mã từ bảng giá --":
                chosen_code = selected_option.split(" - ")[0].strip()
                if chosen_code in dict_mapping_price:
                    info = dict_mapping_price[chosen_code]
                    edited_manual_df.at[idx, 'Mã SP'] = chosen_code
                    edited_manual_df.at[idx, 'Tên chuẩn'] = info['Tên chuẩn']
                    edited_manual_df.at[idx, 'Đơn giá'] = info['Đơn giá']
                    edited_manual_df.at[idx, 'ĐVT'] = info['ĐVT']
                    edited_manual_df.at[idx, 'XUẤT XỨ'] = info['XUẤT XỨ']

        st.session_state.df_manual_edited = edited_manual_df
    else:
        st.success("🎉 Tuyệt vời! Tất cả các mục trong BOQ đã tự động nhận diện hoàn hảo.")
        st.session_state.df_manual_edited = None

    st.markdown("---")
    if st.button("📥 Xuất File Báo Giá Hoàn Chỉnh", type="primary"):
        try:
            with st.spinner("Đang tổng hợp dữ liệu và điền vào mẫu báo giá QN Tech..."):
                
                final_rows = [df_auto]
                if st.session_state.df_manual_edited is not None:
                    valid_manual = st.session_state.df_manual_edited[st.session_state.df_manual_edited['Bỏ qua'] == False].copy()
                    if not valid_manual.empty:
                        valid_manual = valid_manual.drop(columns=['Bỏ qua', 'Chọn Mã Từ Bảng Giá'], errors='ignore')
                        valid_manual['Số lượng'] = pd.to_numeric(valid_manual['Số lượng'], errors='coerce').fillna(0)
                        valid_manual['Đơn giá'] = pd.to_numeric(valid_manual['Đơn giá'], errors='coerce').fillna(0)
                        final_rows.append(valid_manual)

                df_final = pd.concat(final_rows, ignore_index=True)
                df_final['Thành tiền'] = df_final['Số lượng'] * df_final['Đơn giá']

                wb = openpyxl.load_workbook(file_template)
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
                    safe_write(current_row, 5, row_data.get('ĐVT', ''))          
                    safe_write(current_row, 6, row_data.get('Số lượng', 0))      
                    safe_write(current_row, 7, row_data.get('Đơn giá', 0))       
                    safe_write(current_row, 8, row_data.get('Thành tiền', 0))    
                    safe_write(current_row, 9, '')                               

                output = BytesIO()
                wb.save(output)
                output.seek(0)

                final_name = output_filename if output_filename.endswith('.xlsx') else output_filename + '.xlsx'

                st.success("✅ Đã tạo file Báo giá QN Tech thành công!")

                st.download_button(
                    label="💾 TẢI XUỐNG FILE BÁO GIÁ EXCEL",
                    data=output,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo file Excel: {e}")
