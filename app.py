import streamlit as st
import pandas as pd
from fuzzywuzzy import process
import openpyxl
from openpyxl.cell.cell import MergedCell
from io import BytesIO
import datetime
import re
import os

# Cấu hình trang
st.set_page_config(page_title="Phần mềm Báo Giá Tự Động QN Tech", layout="wide")
st.title("Hệ thống Báo giá Tự động - QN Tech")

# --- ĐỊNH NGHĨA FILE LƯU SẴN (CẦN ĐẶT CÙNG THƯ MỤC VỚI APP.PY) ---
DATA_FILE_PATH = "BangGia.xlsx"
TEMPLATE_FILE_PATH = "MauBaoGia.xlsx"

# Kiểm tra xem file lưu sẵn có tồn tại không
file_data = DATA_FILE_PATH if os.path.exists(DATA_FILE_PATH) else None
file_template = TEMPLATE_FILE_PATH if os.path.exists(TEMPLATE_FILE_PATH) else None

# 1. KHU VỰC UPLOAD FILE TỪ KHÁCH HÀNG (GIAO DIỆN)
st.sidebar.header("1. Tải tài liệu lên")
file_boq = st.sidebar.file_uploader("File BOQ từ Khách hàng (Excel)", type=["xlsx"])

# Thông báo trạng thái file hệ thống
if file_data and file_template:
    st.sidebar.success("✅ Đã tải sẵn Bảng giá và Mẫu báo giá hệ thống.")
else:
    st.sidebar.error("❌ Không tìm thấy file 'BangGia.xlsx' hoặc 'MauBaoGia.xlsx' trên hệ thống. Vui lòng kiểm tra lại!")

# Tự động đọc danh sách các hãng từ các sheet trong file Bảng Giá lưu sẵn
selected_brand_sheet = None
if file_data is not None:
    try:
        xls_data = pd.ExcelFile(file_data)
        valid_sheets = []
        for s in xls_data.sheet_names:
            temp_df = pd.read_excel(file_data, sheet_name=s)
            if not temp_df.empty and len(temp_df.columns) > 1:
                valid_sheets.append(s)

        if valid_sheets:
            selected_brand_sheet = st.sidebar.selectbox("Chọn Hãng / Sheet Báo Giá", valid_sheets)
        else:
            st.sidebar.error("File Bảng giá không có sheet dữ liệu hợp lệ!")
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file Bảng giá: {e}")

# Tự động đọc danh sách sheet BOQ thiết bị
selected_boq_sheet = None
if file_boq is not None:
    try:
        xls_boq = pd.ExcelFile(file_boq)
        sheet_names = xls_boq.sheet_names
        default_sheet_idx = 0
        for i, name in enumerate(sheet_names):
            if 'tên thiết bị' in name.lower() or 'thiết bị' in name.lower():
                default_sheet_idx = i
                break
        selected_boq_sheet = st.sidebar.selectbox("Chọn Sheet chứa BOQ thiết bị", sheet_names, index=default_sheet_idx)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file BOQ: {e}")

# 2. KHU VỰC NHẬP THÔNG TIN KHÁCH HÀNG & DỰ ÁN
st.sidebar.header("2. Thông tin Báo giá & Khách hàng")
to_company = st.sidebar.text_input("To: Tên Công ty / Khách hàng", value="CÔNG TY CỔ PHẦN ĐẦU TƯ XÂY DỰNG BÀN THẠCH")
add_company = st.sidebar.text_input("Add: Địa chỉ khách hàng", value="415/17 Trường Chinh, Tân Bình, Hồ Chí Minh")
att_contact = st.sidebar.text_input("Att: Người liên hệ", value="Mr Kiên")
mobile_contact = st.sidebar.text_input("Mobile: Số điện thoại", value="0909123456")
email_contact = st.sidebar.text_input("Email", value="kien@banthach.vn")
ref_content = st.sidebar.text_input("Ref: Nội dung / Tiêu đề báo giá", value="BÁO GIÁ THIẾT BỊ HỆ THỐNG PCCC")

st.sidebar.header("3. Thông tin Quản lý Báo giá")
no_quote = st.sidebar.text_input("No: Số báo giá", value="Q2026/2108/PCCC/BT")
date_quote = st.sidebar.text_input("Date: Ngày báo giá", value="21/08/2026")
validity_quote = st.sidebar.text_input("Validity: Hiệu lực báo giá", value="20/09/2026")
project_name = st.sidebar.text_input("Project: Tên Dự án", value="XƯỞNG SX TỦ LẠNH")

# 4. Ô NHẬP TÊN FILE KHI XUẤT RA
st.sidebar.header("4. Tùy chỉnh tên File xuất")
output_filename = st.sidebar.text_input("Nhập tên file Excel xuất ra:", value="BaoGia_HoanChinh.xlsx")

# Nút bấm chạy chương trình
if st.sidebar.button("Xử lý và Xuất Báo Giá"):
    if not (file_data and file_boq and file_template and selected_brand_sheet):
        st.error("Vui lòng tải lên file BOQ và đảm bảo hệ thống đã nhận diện được Bảng Giá & Mẫu Báo Giá!")
    else:
        try:
            with st.spinner(f"Đang đọc bảng giá hãng [{selected_brand_sheet}] và quy chuẩn mã thiết bị..."):
                # Đọc file Data giá theo sheet hãng đã chọn
                df_data = pd.read_excel(file_data, sheet_name=selected_brand_sheet)
                df_data.columns = [str(c).strip() for c in df_data.columns]

                data_name_col = None
                for c in ['Tên chuẩn', 'TÊN CHUẨN', 'Tên thiết bị', 'TÊN THIẾT BỊ', 'Tên hàng', 'TÊN HÀNG', 'Tên vật tư', 'TÊN VẬT TƯ']:
                    if c in df_data.columns:
                        data_name_col = c
                        break

                if not data_name_col:
                    st.error(f"Sheet '{selected_brand_sheet}' không tìm thấy cột tên thiết bị/tên chuẩn! Các cột hiện có: {list(df_data.columns)}")
                    st.stop()

                if data_name_col != 'Tên chuẩn':
                    df_data = df_data.rename(columns={data_name_col: 'Tên chuẩn'})

                df_data['Tên chuẩn'] = df_data['Tên chuẩn'].astype(str).str.strip()
                df_data = df_data.dropna(subset=['Tên chuẩn']).copy()

                # Đọc file BOQ theo sheet đã chọn
                df_boq = pd.read_excel(file_boq, sheet_name=selected_boq_sheet)
                df_boq.columns = [str(c).strip() for c in df_boq.columns]

                name_col_found = None
                possible_name_cols = ['Tên thiết bị BOQ', 'Diễn giải', 'DIỄN GIẢI', 'Tên hàng', 'TÊN HÀNG', 'Tên vật tư', 'TÊN VẬT TƯ', 'Tên thiết bị', 'TÊN THIẾT BỊ', 'Tên công tác', 'TÊN CÔNG TÁC']
                for c in possible_name_cols:
                    if c in df_boq.columns:
                        name_col_found = c
                        break

                if not name_col_found and len(df_boq.columns) > 1:
                    name_col_found = df_boq.columns[1]

                if not name_col_found:
                    st.error(f"Không nhận diện được cột tên thiết bị trong file BOQ! Các cột hiện có: {list(df_boq.columns)}")
                    st.stop()

                if name_col_found != 'Tên thiết bị BOQ':
                    df_boq = df_boq.rename(columns={name_col_found: 'Tên thiết bị BOQ'})

                qty_col_found = None
                possible_qty_cols = ['Khối lượng', 'Số lượng', 'KHỐI LƯỢNG', 'SỐ LƯỢNG', 'SL', 'Sl', 'sl']
                for c in possible_qty_cols:
                    if c in df_boq.columns:
                        qty_col_found = c
                        break

                if not qty_col_found and len(df_boq.columns) > 3:
                    qty_col_found = df_boq.columns[3]

                if qty_col_found and qty_col_found != 'Số lượng':
                    df_boq = df_boq.rename(columns={qty_col_found: 'Số lượng'})
                elif not qty_col_found:
                    df_boq['Số lượng'] = 0

                df_boq = df_boq.dropna(subset=['Tên thiết bị BOQ']).copy()

                df_boq['Tên thiết bị BOQ'] = df_boq['Tên thiết bị BOQ'].astype(str).str.replace('\u3000', ' ').str.strip()
                df_boq['Số lượng'] = pd.to_numeric(df_boq['Số lượng'], errors='coerce').fillna(0)

                # Dùng sort=False để giữ nguyên thứ tự xuất hiện ban đầu từ BOQ gốc
                df_boq = df_boq.groupby('Tên thiết bị BOQ', sort=False, as_index=False)['Số lượng'].sum()

                # Hàm bóc tách đường kính (DN) từ tên thiết bị BOQ
                def extract_dn(text):
                    match = re.search(r'DN\s*(\d+)', str(text), re.IGNORECASE)
                    if match:
                        return int(match.group(1))
                    return 0

                # Hàm chuẩn hóa tên thiết bị tích hợp toàn bộ quy tắc nghiệp vụ & Hochiki
                def map_ten_va_quy_tac(ten_boq):
                    text_lower = str(ten_boq).lower()
                    dn_val = extract_dn(ten_boq)

                    # 1. Nhóm thiết bị báo cháy Hochiki thực tế
                    if 'khói' in text_lower and ('địa chỉ' in text_lower or 'đ/c' in text_lower):
                        return "Đầu báo khói quang địa chỉ kèm đế"
                    if 'khói' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower):
                        return "Đầu báo khói quang loại thường kèm đế"
                    if 'nhiệt' in text_lower and ('cổ điển' in text_lower or 'thường' in text_lower):
                        return "Đầu nhiệt tích hợp loại thường kèm đế"
                    if 'giám sát đầu báo thường' in text_lower or '(zm)' in text_lower:
                        return "Module 2 ngõ vào giám sát đầu báo thường"
                    if 'nút nhấn' in text_lower and 'địa chỉ' in text_lower:
                        return "Nút nhấn báo cháy địa chỉ"
                    if ('còi đèn' in text_lower or 'chuông đèn' in text_lower) and 'địa chỉ' in text_lower:
                        return "Còi đèn chớp báo cháy địa chỉ kèm đế"
                    if 'điều khiển' in text_lower and '(cm)' in text_lower:
                        return "Module 1 ngõ ra Rơ le"
                    if 'giám sát' in text_lower and '(mm)' in text_lower:
                        return "Module 1 ngõ vào giám sát"

                    # 2. Tránh nhầm lẫn van bi với van cổng
                    if 'van bi' in text_lower:
                        return ten_boq

                    # 3. Lọc rác chữ Y -> Y Lọc
                    if 'lọc rác chữ y' in text_lower or 'y lọc' in text_lower:
                        return str(ten_boq).replace('Lọc rác chữ Y', 'Y lọc').replace('lọc rác chữ y', 'Y lọc')

                    # 4. Công tắc áp lực -> Công tắc áp suất
                    if 'công tắc áp lực' in text_lower:
                        return str(ten_boq).replace('áp lực', 'áp suất')

                    # 5. Khớp nối mềm chống rung bằng kim loại / Chống rung -> Khớp nối mềm
                    if 'chống rung' in text_lower or 'khớp nối mềm' in text_lower:
                        if dn_val > 0:
                            return f"Khớp nối mềm DN{dn_val}"
                        return "Khớp nối mềm"

                    # 6. Đồng hồ áp lực -> Đồng hồ áp suất
                    if 'đồng hồ áp lực' in text_lower:
                        return str(ten_boq).replace('áp lực', 'áp suất')

                    # 7. Van 1 chiều: <= DN50 là đồng, > DN50 là gang
                    if 'van 1 chiều' in text_lower or 'van một chiều' in text_lower or 'check valve' in text_lower:
                        if 0 < dn_val <= 50:
                            return f"Van một chiều đồng DN{dn_val}"
                        elif dn_val > 50:
                            return f"Van một chiều gang DN{dn_val}"

                    # 8. Van chặn / Van cổng (có hoặc không có kèm giám sát / công tắc giám sát)
                    if 'van chặn' in text_lower or 'van cổng' in text_lower or 'gate valve' in text_lower:
                        has_giamsat = ('giám sát' in text_lower or 'kèm giám sát' in text_lower or 'công tắc giám sát' in text_lower)
                        if 'os&y' in text_lower or 'osy' in text_lower:
                            if has_giamsat:
                                return f"Van Cổng OS&Y Kèm Giám Sát DN{dn_val}" if dn_val > 0 else "Van Cổng OS&Y Kèm Giám Sát"
                            else:
                                return f"Van cổng OS&Y DN{dn_val}" if dn_val > 0 else "Van cổng OS&Y"
                        else:
                            if has_giamsat:
                                return f"Van cổng NRS DN{dn_val} kèm công tắc giám sát" if dn_val > 0 else "Van cổng NRS kèm công tắc giám sát"
                            else:
                                if 0 < dn_val <= 50:
                                    return f"Van chặn DN{dn_val}"
                                elif dn_val > 50:
                                    return f"Van cổng NRS DN{dn_val}"

                    return ten_boq

                df_boq['Tên BOQ Chuẩn Hóa'] = df_boq['Tên thiết bị BOQ'].apply(map_ten_va_quy_tac)

                danh_sach_chuan = df_data['Tên chuẩn'].tolist()

                def tim_ma_phu_hop(ten_chuan_boq, ten_goc_boq):
                    dn_val = extract_dn(ten_goc_boq)
                    text_lower = str(ten_goc_boq).lower()
                    item_lower_check = str(ten_chuan_boq).lower()

                    # Ưu tiên khớp chính xác mã Hochiki thực tế
                    if 'đầu báo khói quang địa chỉ' in item_lower_check:
                        for item in danh_sach_chuan:
                            if 'aln-en' in str(item).lower():
                                return item
                    if 'module 2 ngõ vào giám sát đầu báo thường' in item_lower_check or 'zm' in text_lower:
                        for item in danh_sach_chuan:
                            if 'chq-dzm' in str(item).lower():
                                return item
                    if 'nút nhấn báo cháy địa chỉ' in item_lower_check:
                        for item in danh_sach_chuan:
                            if 'hcp-e' in str(item).lower():
                                return item
                    if 'còi đèn chớp báo cháy địa chỉ' in item_lower_check:
                        for item in danh_sach_chuan:
                            if 'chq-wsb2' in str(item).lower():
                                return item

                    # Không xử lý nhầm van bi
                    if 'van bi' in text_lower:
                        match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                        return match if score > 70 else "Cần check thủ công"

                    # 1. Quy tắc Van chặn / Van cổng (có hoặc không có giám sát)
                    if 'van chặn' in text_lower or 'van cổng' in text_lower or 'gate valve' in text_lower:
                        has_giamsat = ('giám sát' in text_lower or 'kèm giám sát' in text_lower or 'công tắc giám sát' in text_lower)
                        if 'os&y' in text_lower or 'osy' in text_lower:
                            for item in danh_sach_chuan:
                                item_str = str(item).lower()
                                if str(dn_val) in item_str and 'os&y' in item_str:
                                    if has_giamsat and ('kèm giám sát' in item_str or 'giám sát' in item_str):
                                        return item
                                    elif not has_giamsat and not ('kèm giám sát' in item_str or 'giám sát' in item_str):
                                        return item
                        else:
                            if 0 < dn_val <= 50 and not has_giamsat:
                                for item in danh_sach_chuan:
                                    if 'z15w' in str(item).lower() and str(dn_val) in str(item):
                                        return item
                            else:
                                for item in danh_sach_chuan:
                                    item_str = str(item).lower()
                                    if str(dn_val) in item_str and 'nrs' in item_str:
                                        if has_giamsat and ('kèm giám sát' in item_str or 'công tắc giám sát' in item_str):
                                            return item
                                        elif not has_giamsat and not ('kèm giám sát' in item_str or 'công tắc giám sát' in item_str):
                                            return item

                    # 2. Quy tắc Van 1 chiều (Đồng <= DN50 -> H14W-16T, Gang > DN50 -> H44X-16Q)
                    if 'van 1 chiều' in text_lower or 'van một chiều' in text_lower or 'check valve' in text_lower:
                        for item in danh_sach_chuan:
                            if 0 < dn_val <= 50 and ('h14w' in str(item).lower() or 'đồng' in str(item).lower()) and str(dn_val) in str(item):
                                return item
                            elif dn_val > 50 and ('h44x' in str(item).lower() or 'gang' in str(item).lower()) and str(dn_val) in str(item):
                                return item

                    # Khớp thông thường bằng FuzzyWuzzy cho các thiết bị khác
                    match, score = process.extractOne(str(ten_chuan_boq), danh_sach_chuan)
                    return match if score > 70 else "Cần check thủ công"

                df_boq['Tên Map'] = df_boq.apply(lambda row: tim_ma_phu_hop(row['Tên BOQ Chuẩn Hóa'], row['Tên thiết bị BOQ']), axis=1)
                df_result = pd.merge(df_boq, df_data, left_on='Tên Map', right_on='Tên chuẩn', how='left')

                df_result['Đơn giá'] = pd.to_numeric(df_result['Đơn giá'], errors='coerce').fillna(0)
                df_result['Thành tiền'] = df_result['Số lượng'] * df_result['Đơn giá']

                st.subheader(f"Bảng dữ liệu sau khi quy chuẩn và khớp với hãng: {selected_brand_sheet}")
                cols_to_show = [c for c in ['Tên thiết bị BOQ', 'Tên Map', 'Mã SP', 'XUẤT XỨ', 'ĐVT', 'Số lượng', 'Đơn giá', 'Thành tiền'] if c in df_result.columns]
                st.dataframe(df_result[cols_to_show])

                # --- GHI VÀO FILE TEMPLATE EXCEL ĐÚNG TỌA ĐỘ ---
                wb = openpyxl.load_workbook(file_template)
                ws = wb.active

                def safe_write(row, column, val):
                    cell = ws.cell(row=row, column=column)
                    if isinstance(cell, MergedCell):
                        return
                    cell.value = val

                safe_write(row=6,  column=2, val=to_company)       
                safe_write(row=7,  column=2, val=add_company)      
                safe_write(row=8,  column=2, val=att_contact)      
                safe_write(row=9,  column=2, val=mobile_contact)   
                safe_write(row=10, column=2, val=email_contact)    
                safe_write(row=11, column=2, val=ref_content)      

                safe_write(row=6, column=7, val=no_quote)          
                safe_write(row=7, column=7, val=date_quote)        
                safe_write(row=9, column=7, val=validity_quote)    
                safe_write(row=10, column=7, val=project_name)   

                start_row = 14
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

                st.success(f"Tạo file báo giá hãng [{selected_brand_sheet}] thành công!")

                st.download_button(
                    label="📥 Tải Báo Giá Hoàn Chỉnh (Excel)",
                    data=output,
                    file_name=final_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Có lỗi xảy ra trong quá trình xử lý: {e}")
