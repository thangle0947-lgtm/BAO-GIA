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

# ==========================================
# XỬ LÝ DỮ LIỆU CHÍNH
# ==========================================
if st.sidebar.button("Xử lý và Xuất Báo Giá", type="primary"):
    if not (file_data and file_boq and file_template and selected_brand_sheet):
        st.error("⚠️ Vui lòng tải lên file BOQ và kiểm tra đầy đủ Bảng Giá / Mẫu Báo Giá!")
    else:
        try:
            with st.spinner("Đang phân tích BOQ, map mã thiết bị và căn chỉnh file Excel..."):
                
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

                # --- HÀM HỖ TRỢ XỬ LÝ CHUỖI VÀ MAP DATA ---
                def extract_dn(text):
                    match = re.search(r'DN\s*(\d+)', str(text), re.IGNORECASE)
                    return int(match.group(1)) if match else 0

                def map_ten_va_quy_tac(ten_boq):
                    text_lower = str(ten_boq).lower()
                    dn_val = extract_dn(ten_boq)

                    if 'khói' in text_lower and ('địa chỉ' in text_lower or 'đ/c' in text_lower): return "Đầu báo khói quang địa chỉ kèm đế"
                    if 'khói' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower): return "Đầu báo khói quang loại thường kèm đế"
                    if 'nhiệt' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower): return "Đầu nhiệt tích hợp loại thường kèm đế"
                    if 'giám sát đầu báo thường' in text_lower or '(zm)' in text_lower: return "Module 2 ngõ vào giám sát đầu báo thường"
                    if 'nút nhấn' in text_lower and 'địa chỉ' in text_lower: return "Nút nhấn báo cháy địa chỉ"
                    if ('còi đèn' in text_lower or 'chuông đèn' in text_lower) and 'địa chỉ' in text_lower: return "Còi đèn chớp báo cháy địa chỉ kèm đế"
                    if 'điều khiển' in text_lower and '(cm)' in text_lower: return "Module 1 ngõ ra Rơ le"
                    if 'giám sát' in text_lower and '(mm)' in text_lower: return "Module 1 ngõ vào giám sát"
                    if 'lọc rác chữ y' in text_lower or 'y lọc' in text_lower: return str(ten_boq).replace('Lọc rác chữ Y', 'Y lọc').replace('lọc rác chữ y', 'Y lọc')
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
                df_result['Thành tiền'] = df_result['Số lượng'] * df_result['Đơn giá']

                st.subheader("Bảng xem trước dữ liệu Báo giá:")
                st.dataframe(df_result[['Tên thiết bị BOQ', 'Tên Map', 'Mã SP', 'Số lượng', 'Đơn giá', 'Thành tiền']])

                # ==========================================
                # XỬ LÝ FILE EXCEL VÀ CĂN CHỈNH BẰNG ẨN DÒNG (AN TOÀN TUYỆT ĐỐI)
                # ==========================================
                wb = openpyxl.load_workbook(file_template)
                ws = wb.active

                def safe_write(row, col, val):
                    cell = ws.cell(row=row, column=col)
                    if not isinstance(cell, MergedCell):
                        cell.value = val

                # 1. Ghi thông tin Header
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

                # 2. Xử lý Ẩn dòng thông minh
                start_row = 14  
                num_items = len(df_result)

                footer_row = start_row + 1
                for r in range(start_row, start_row + 1000):
                    v2 = str(ws.cell(row=r, column=2).value or "").strip().lower()
                    v3 = str(ws.cell(row=r, column=3).value or "").strip().lower()
                    if 'cộng' in v2 or 'cộng' in v3 or 'total' in v2:
                        footer_row = r
                        break

                available_rows = footer_row - start_row

                if num_items <= available_rows:
                    # Ẩn các dòng dư thừa, giữ nguyên vẹn form và chữ ký ở dưới
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

                # 3. Ghi dữ liệu thiết bị
                for index, row_data in df_result.iterrows():
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

                st.success("✅ Đã xử lý xong! File Excel đã được ẩn dòng thừa và giữ nguyên cấu trúc.")

                st.download_button(
                    label="📥 TẢI BÁO GIÁ HOÀN CHỈNH (EXCEL)",
                    data=output,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
