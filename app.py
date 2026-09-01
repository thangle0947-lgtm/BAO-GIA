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
# CẤU HÌNH TRANG & TIÊU ĐỀ
# ==========================================
st.set_page_config(page_title="Phần mềm Báo Giá Tự Động QN Tech", layout="wide")
st.title("Hệ thống Báo giá Tự động - QN Tech")

# ==========================================
# KHAI BÁO ĐƯỜNG DẪN FILE HỆ THỐNG
# ==========================================
DATA_FILE_PATH = "BangGia.xlsx"
TEMPLATE_FILE_PATH = "MauBaoGia.xlsx"

file_data = DATA_FILE_PATH if os.path.exists(DATA_FILE_PATH) else None
file_template = TEMPLATE_FILE_PATH if os.path.exists(TEMPLATE_FILE_PATH) else None

if file_data and file_template:
    st.sidebar.success("✅ Đã nhận diện Bảng giá và Mẫu báo giá trên hệ thống.")
else:
    st.sidebar.error("❌ Không tìm thấy file 'BangGia.xlsx' hoặc 'MauBaoGia.xlsx'. Vui lòng kiểm tra lại!")

# ==========================================
# GIAO DIỆN BÊN TRÁI (SIDEBAR) - NHẬP LIỆU
# ==========================================
st.sidebar.header("1. Tải tài liệu lên")
file_boq = st.sidebar.file_uploader("File BOQ từ Khách hàng (Excel)", type=["xlsx"])

selected_brand_sheet = None
if file_data is not None:
    try:
        xls_data = pd.ExcelFile(file_data)
        valid_sheets = [s for s in xls_data.sheet_names if not pd.read_excel(file_data, sheet_name=s).empty]
        if valid_sheets:
            selected_brand_sheet = st.sidebar.selectbox("Chọn Hãng / Sheet Bảng Giá", valid_sheets)
        else:
            st.sidebar.error("File Bảng giá trống hoặc không hợp lệ!")
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

st.sidebar.header("2. Thông tin Báo giá & Khách hàng")
to_company = st.sidebar.text_input("To: Tên Công ty / Khách hàng", value="CÔNG TY CỔ PHẦN ĐẦU TƯ PHÁT TRIỂN MINH GLOBAL")
add_company = st.sidebar.text_input("Add: Địa chỉ khách hàng", value="415/17 Trường Chinh, Tân Bình, Hồ Chí Minh")
att_contact = st.sidebar.text_input("Att: Người liên hệ", value="Mr Kiên")
mobile_contact = st.sidebar.text_input("Mobile: Số điện thoại", value="0909123456")
email_contact = st.sidebar.text_input("Email", value="kien@minhglobal.vn")
ref_content = st.sidebar.text_input("Ref: Nội dung / Tiêu đề", value="BÁO GIÁ THIẾT BỊ HỆ THỐNG PCCC")

st.sidebar.header("3. Thông tin Quản lý Báo giá")
no_quote = st.sidebar.text_input("No: Số báo giá", value="Q2026/2108/PCCC/MG")
date_quote = st.sidebar.text_input("Date: Ngày báo giá", value="01/09/2026")
validity_quote = st.sidebar.text_input("Validity: Hiệu lực báo giá", value="30/09/2026")
project_name = st.sidebar.text_input("Project: Tên Dự án", value="NHÀ MÁY TRẦN ĐỨC / BỆNH VIỆN QUÂN Y 7B")

st.sidebar.header("4. Tùy chỉnh tên File xuất")
output_filename = st.sidebar.text_input("Nhập tên file Excel xuất ra:", value="BaoGia_HoanChinh.xlsx")

# Khởi tạo Session State
if 'df_raw_result' not in st.session_state:
    st.session_state.df_raw_result = None
if 'df_manual_edited' not in st.session_state:
    st.session_state.df_manual_edited = None

# ==========================================
# XỬ LÝ DỮ LIỆU BAN ĐẦU
# ==========================================
if st.sidebar.button("Xử lý và Phân Tích BOQ", type="primary"):
    if not (file_data and file_boq and file_template and selected_brand_sheet):
        st.error("⚠️ Vui lòng tải lên file BOQ và kiểm tra đầy đủ Bảng Giá / Mẫu Báo Giá!")
    else:
        try:
            with st.spinner("Đang phân tích BOQ và map mã thiết bị..."):
                
                # --- ĐỌC DỮ LIỆU BẢNG GIÁ ---
                df_data = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
                df_data.columns = [str(c).strip() for c in df_data.columns]
                
                data_name_col = next((c for c in df_data.columns if c.lower() in ['tên chuẩn', 'tên thiết bị', 'tên hàng', 'tên vật tư', 'diễn giải']), None)
                if not data_name_col:
                    st.error(f"Sheet '{selected_brand_sheet}' không tìm thấy cột Tên chuẩn!")
                    st.stop()
                
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

                # --- HÀM HỖ TRỢ XỬ LÝ CHUỖI VÀ TỰ ĐỘNG NHẬN DIỆN CÁP ---
                def extract_dn(text):
                    match = re.search(r'DN\s*(\d+)', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 0

                def map_ten_va_quy_tac(ten_boq):
                    text_lower = str(ten_boq).lower()
                    dn_val = extract_dn(ten_boq)

                    if 'silicone' in text_lower or 'chống cháy' in text_lower or 'cap' in text_lower or 'cáp' in text_lower:
                        has_khong_nhieu = 'không chống nhiễu' in text_lower or 'ko chống nhiễu' in text_lower
                        has_chong_nhieu = ('chống nhiễu' in text_lower or 'shield' in text_lower or 'st' in text_lower) and not has_khong_nhieu
                        shield_text = "chống cháy chống nhiễu" if has_chong_nhieu else "chống cháy không chống nhiễu"

                        core_match = re.search(r'(\d+)\s*(?:lõi|c\b|core|x)', text_lower)
                        num_cores = int(core_match.group(1)) if core_match else 2

                        size_val = "1.5"
                        if '250' in text_lower or '2.5' in text_lower or '2,5' in text_lower or '25' in text_lower:
                            size_val = "2.5"
                        elif '1.5' in text_lower or '1,5' in text_lower or '15' in text_lower:
                            size_val = "1.5"
                        elif '1.0' in text_lower or '1,0' in text_lower:
                            size_val = "1.0"
                        elif '0.5' in text_lower or '0,5' in text_lower:
                            size_val = "0.50"
                        elif '0.75' in text_lower or '0,75' in text_lower:
                            size_val = "0.75"
                        elif '4.0' in text_lower or '4,0' in text_lower:
                            size_val = "4.0"
                        else:
                            floats = re.findall(r'(\d+[\.,]\d+|\d+)', text_lower)
                            for f in floats:
                                f_clean = f.replace(',', '.')
                                try:
                                    val_f = float(f_clean)
                                    if val_f in [0.5, 0.50, 0.75, 1.0, 1.25, 1.5, 2.0, 2.5, 4.0]:
                                        size_val = f_clean
                                        break
                                except:
                                    pass
                        
                        try:
                            f_val = float(size_val)
                            if f_val == 0.5: size_str = "0.50mm2"
                            elif f_val == 0.75: size_str = "0.75mm2"
                            elif f_val == 1.0 or f_val == 1: size_str = "1.0mm2"
                            elif f_val == 1.25: size_str = "1.25mm2"
                            elif f_val == 1.5: size_str = "1.5mm2"
                            elif f_val == 2.0 or f_val == 2: size_str = "2.0mm2"
                            elif f_val == 2.5: size_str = "2.5mm2"
                            elif f_val == 4.0 or f_val == 4: size_str = "4.0mm2"
                            else: size_str = f"{f_val:.2f}mm2"
                        except:
                            size_str = "1.5mm2"

                        return f"Cáp tín hiệu điều khiển Silicone {shield_text}, {num_cores} lõi, {size_str}"

                    if 'khói' in text_lower and ('địa chỉ' in text_lower or 'đ/c' in text_lower): return "Đầu báo khói quang địa chỉ kèm đế"
                    if 'khói' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower): return "Đầu báo khói quang loại thường kèm đế"
                    if 'nhiệt' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower): return "Đầu nhiệt tích hợp loại thường kèm đế"
                    if 'giám sát đầu báo thường' in text_lower or '(zm)' in text_lower: return "Module 2 ngõ vào giám sát đầu báo thường"
                    if 'nút nhấn' in text_lower and 'địa chỉ' in text_lower: return "Nút nhấn báo cháy địa chỉ"
                    if ('còi đèn' in text_lower or 'chuông đèn' in text_lower) and 'địa chỉ' in text_lower: return "Còi đèn chớp báo cháy địa chỉ kèm đế"
                    if 'điều khiển' in text_lower and '(cm)' in text_lower: return "Module 1 ngõ ra Rơ le"
                    if 'giám sát' in text_lower and '(mm)' in text_lower: return "Module 1 ngõ vào giám sát"
                    if 'lọc rác chữ y' in text_lower or 'y lọc' in text_lower: return str(ten_boq).replace('Lọc rác chữ Y', 'Y lọc')
                    if 'công tắc áp lực' in text_lower or 'đồng hồ áp lực' in text_lower: return str(ten_boq).replace('áp lực', 'áp suất')
                    if 'chống rung' in text_lower or 'khớp nối mềm' in text_lower: return f"Khớp nối mềm DN{dn_val}" if dn_val > 0 else "Khớp nối mềm"

                    if 'van 1 chiều' in text_lower or 'van một chiều' in text_lower or 'check valve' in text_lower:
                        if 0 < dn_val <= 50: return f"Van một chiều đồng DN{dn_val}"
                        elif dn_val > 50: return f"Van một chiều gang DN{dn_val}"

                    if 'van chặn' in text_lower or 'van cổng' in text_lower or 'gate valve' in text_lower:
                        has_giamsat = ('giám sát' in text_lower or 'kèm giám sát' in text_lower or 'công tắc giám sát' in text_lower)
                        if 'os&y' in text_lower or 'osy' in text_lower:
                            return f"Van Cổng OS&Y Kèm Giám Sát DN{dn_val}" if (dn_val > 0 and has_giamsat) else (f"Van cổng OS&Y DN{dn_val}" if dn_val > 0 else "Van cổng OS&Y")
                        else:
                            if has_giamsat: return f"Van cổng NRS DN{dn_val} kèm công tắc giám sát" if dn_val > 0 else "Van cổng NRS kèm công tắc giám sát"
                            else: return f"Van chặn DN{dn_val}" if 0 < dn_val <= 50 else f"Van cổng NRS DN{dn_val}"
                    return ten_boq

                def tim_ma_phu_hop(ten_chuan_boq, ten_goc_boq):
                    dn_val = extract_dn(ten_goc_boq)
                    text_lower = str(ten_goc_boq).lower()
                    item_lower_check = str(ten_chuan_boq).lower()

                    if 'silicone' in item_lower_check or 'chống cháy' in item_lower_check or 'cáp' in item_lower_check:
                        is_shielded_item = ('chống nhiễu' in item_lower_check and 'không chống nhiễu' not in item_lower_check) or 'cszs' in item_lower_check
                        size_match = re.search(r'(\d+[\.,]\d+|\d+)mm2', item_lower_check)
                        target_size = size_match.group(1) if size_match else "1.5"
                        core_match = re.search(r'(\d+)\s*lõi', item_lower_check)
                        target_core = core_match.group(1) if core_match else "2"

                        for item in danh_sach_chuan:
                            item_str = str(item).lower()
                            item_is_shielded = ('chống nhiễu' in item_str and 'không chống nhiễu' not in item_str) or 'cszs' in item_str
                            if item_is_shielded == is_shielded_item:
                                if target_size in item_str and (f"{target_core} lõi" in item_str or f"{target_core}lõi" in item_str):
                                    return item
                        return "Cần check thủ công"

                    if 'đầu báo khói quang địa chỉ' in item_lower_check: return next((item for item in danh_sach_chuan if 'aln-en' in str(item).lower()), "Cần check thủ công")
                    if 'module 2 ngõ vào giám sát đầu báo thường' in item_lower_check or 'zm' in text_lower: return next((item for item in danh_sach_chuan if 'chq-dzm' in str(item).lower()), "Cần check thủ công")
                    if 'nút nhấn báo cháy địa chỉ' in item_lower_check: return next((item for item in danh_sach_chuan if 'hcp-e' in str(item).lower()), "Cần check thủ công")
                    if 'còi đèn chớp báo cháy địa chỉ' in item_lower_check: return next((item for item in danh_sach_chuan if 'chq-wsb2' in str(item).lower()), "Cần check thủ công")

                    if 'van bi' in text_lower:
                        match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                        return match if score > 70 else "Cần check thủ công"

                    if 'van cổng' in text_lower or 'van chặn' in text_lower:
                        has_giamsat = ('giám sát' in text_lower)
                        for item in danh_sach_chuan:
                            item_str = str(item).lower()
                            if str(dn_val) in item_str:
                                if 'os&y' in text_lower or 'osy' in text_lower:
                                    if 'os&y' in item_str and (has_giamsat == ('giám sát' in item_str)): return item
                                else:
                                    if 0 < dn_val <= 50 and not has_giamsat and 'z15w' in item_str: return item
                                    if 'nrs' in item_str and (has_giamsat == ('giám sát' in item_str)): return item

                    match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                    return match if score > 70 else "Cần check thủ công"

                df_boq['Tên BOQ Chuẩn Hóa'] = df_boq['Tên thiết bị BOQ'].apply(map_ten_va_quy_tac)
                df_boq['Tên Map'] = df_boq.apply(lambda row: tim_ma_phu_hop(row['Tên BOQ Chuẩn Hóa'], row['Tên thiết bị BOQ']), axis=1)
                
                df_result = pd.merge(df_boq, df_data, left_on='Tên Map', right_on='Tên chuẩn', how='left')
                df_result['Đơn giá'] = pd.to_numeric(df_result['Đơn giá'], errors='coerce').fillna(0)
                
                # Lưu vào session state
                st.session_state.df_raw_result = df_result
                st.success("✅ Phân tích BOQ thành công! Vui lòng kiểm tra danh sách bên dưới.")

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")

# ==========================================
# GIAO DIỆN KIỂM TRA & NHẬP THỦ CÔNG
# ==========================================
if st.session_state.df_raw_result is not None:
    df_current = st.session_state.df_raw_result
    
    st.subheader("📋 Bảng Kết Quả & Xử Lý Mã Cần Check Thủ Công")
    st.info("💡 Với những dòng 'Cần check thủ công', bạn có thể nhập trực tiếp Mã SP, Tên chuẩn, Xuất xứ, ĐVT, Đơn giá vào bảng dưới đây. Nếu muốn loại bỏ thiết bị đó ra khỏi báo giá, hãy tích chọn ô **Bỏ qua**.")

    # Tách thành 2 phần: Đã map thành công và Cần check thủ công
    mask_manual = (df_current['Tên Map'] == 'Cần check thủ công') | (df_current['Mã SP'].isna()) | (df_current['Mã SP'] == '')
    df_auto = df_current[~mask_manual].copy()
    df_manual = df_current[mask_manual].copy()

    if not df_manual.empty:
        # Chuẩn bị bảng chỉnh sửa thủ công
        if 'Bỏ qua' not in df_manual.columns:
            df_manual.insert(0, 'Bỏ qua', False)
        if 'XUẤT XỨ' not in df_manual.columns:
            df_manual['XUẤT XỨ'] = 'LS/ VN'
        if 'ĐVT' not in df_manual.columns:
            df_manual['ĐVT'] = 'Mét' if 'cáp' in df_manual['Tên thiết bị BOQ'].str.lower().to_string() else 'Cái'
        if 'Tên chuẩn' not in df_manual.columns or df_manual['Tên chuẩn'].isna().any():
            df_manual['Tên chuẩn'] = df_manual['Tên thiết bị BOQ']

        # Hiển thị trình chỉnh sửa dữ liệu (Data Editor)
        edited_manual_df = st.data_editor(
            df_manual[['Bỏ qua', 'Tên thiết bị BOQ', 'Mã SP', 'Tên chuẩn', 'XUẤT XỨ', 'ĐVT', 'Số lượng', 'Đơn giá']],
            num_rows="fixed",
            key="manual_editor",
            use_container_width=True
        )
        st.session_state.df_manual_edited = edited_manual_df
    else:
        st.success("🎉 Tất cả các mục trong BOQ đã được tự động nhận diện thành công!")
        st.session_state.df_manual_edited = None

    # Nút xuất file chính thức
    if st.button("🚀 Xác Nhận & Xuất File Báo Giá Hoàn Chỉnh", type="primary"):
        try:
            with st.spinner("Đang tổng hợp dữ liệu và xuất file Excel..."):
                
                # Gộp dữ liệu tự động và dữ liệu thủ công (đã loại bỏ các dòng bị tích 'Bỏ qua')
                final_rows = [df_auto]
                if st.session_state.df_manual_edited is not None:
                    valid_manual = st.session_state.df_manual_edited[st.session_state.df_manual_edited['Bỏ qua'] == False].copy()
                    if not valid_manual.empty:
                        # Bỏ cột 'Bỏ qua' trước khi merge
                        valid_manual = valid_manual.drop(columns=['Bỏ qua'])
                        # Đảm bảo các cột số liệu đúng kiểu
                        valid_manual['Số lượng'] = pd.to_numeric(valid_manual['Số lượng'], errors='coerce').fillna(0)
                        valid_manual['Đơn giá'] = pd.to_numeric(valid_manual['Đơn giá'], errors='coerce').fillna(0)
                        final_rows.append(valid_manual)

                df_final = pd.concat(final_rows, ignore_index=True)
                df_final['Thành tiền'] = df_final['Số lượng'] * df_final['Đơn giá']

                # --- XỬ LÝ FILE EXCEL MẪU ---
                wb = openpyxl.load_workbook(file_template)
                ws = wb.active

                def safe_write(row, col, val):
                    cell = ws.cell(row=row, column=col)
                    if not isinstance(cell, MergedCell):
                        cell.value = val

                # Ghi thông tin Header
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

                # Xử lý Ẩn/Thêm dòng linh hoạt
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

                # Ghi dữ liệu thiết bị vào file Excel
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

                st.success("✅ Đã xuất file báo giá hoàn chỉnh thành công!")

                st.download_button(
                    label="📥 TẢI BÁO GIÁ HOÀN CHỈNH (EXCEL)",
                    data=output,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
                st.error(f"Đã xảy ra lỗi khi tạo file Excel: {e}")
