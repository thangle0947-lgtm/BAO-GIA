import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime
import copy
import os
import re

# --- CẤU HÌNH TRANG & GIAO DIỆN ---
st.set_page_config(
    page_title="QN TECH - Hệ Thống Báo Giá",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS TÙY CHỈNH (GIAO DIỆN DARK THEME) ---
st.markdown("""
<style>
    .stApp { background-color: #121212; color: #E0E0E0; }
    [data-testid="stSidebar"] { background-color: #1A1A24; border-right: 1px solid #333; }
    h1 { color: #E85D04 !important; font-family: 'Segoe UI', Tahoma, sans-serif; font-weight: 700 !important; border-bottom: 2px solid #0C2340; padding-bottom: 10px; }
    h2, h3 { color: #4A90E2 !important; font-family: 'Segoe UI', Tahoma, sans-serif; }
    .stButton > button[data-baseweb="button"] { background-color: #E85D04; color: white; border-radius: 5px; border: none; font-weight: bold; transition: all 0.3s ease; }
    .stButton > button[data-baseweb="button"]:hover { background-color: #FF7A00; box-shadow: 0px 4px 10px rgba(232, 93, 4, 0.4); }
    .stAlert { background-color: #1E1E2E; border: 1px solid #333; color: #E0E0E0; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; border: 1px solid #333; }
    hr { border-top: 1px solid #333; }
    .streamlit-expanderHeader { background-color: #1A1A24; color: #4A90E2; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
    <h1 style="margin: 0; border: none; font-size: 36px;">
        <span style="color: #4A90E2;">💠 QN TECH</span> <span style="color: #E0E0E0; font-size: 24px; font-weight: 400;">| Hệ Thống Xử Lý Báo Giá Thông Minh</span>
    </h1>
</div>
""", unsafe_allow_html=True)

# --- CẤU HÌNH ĐƯỜNG DẪN TÍCH HỢP SẴN ---
TEMPLATE_PATH = "MauBaoGia.xlsx"
PRICE_LIST_PATH = "BangGia.xlsx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"❌ Không tìm thấy file mẫu `{TEMPLATE_PATH}`.")
    st.stop()

# --- SIDEBAR (GIAO DIỆN TẢI FILE) ---
st.sidebar.markdown("<h2 style='text-align: center; color: #E85D04;'>📥 TẢI DỮ LIỆU</h2>", unsafe_allow_html=True)
uploaded_boq = st.sidebar.file_uploader("Tải lên File BOQ khối lượng gốc (Excel/CSV)", type=["xlsx", "xls", "csv"])

boq_sheet_selected = None
price_sheet_selected = None

if uploaded_boq is not None and uploaded_boq.name.endswith(('xlsx', 'xls')):
    try:
        xl_boq = pd.ExcelFile(uploaded_boq)
        st.sidebar.markdown("---")
        boq_sheet_selected = st.sidebar.selectbox("📂 Chọn Sheet chứa khối lượng (BOQ):", xl_boq.sheet_names)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc file BOQ: {e}")

if os.path.exists(PRICE_LIST_PATH):
    st.sidebar.markdown("---")
    try:
        xl_price = pd.ExcelFile(PRICE_LIST_PATH)
        price_sheet_selected = st.sidebar.selectbox("📋 Chọn Sheet Bảng Giá áp dụng:", xl_price.sheet_names)
    except Exception as e:
        st.sidebar.error(f"Lỗi đọc Bảng Giá: {e}")

if "df_editable" not in st.session_state: st.session_state.df_editable = None
if "price_list_cache" not in st.session_state: st.session_state.price_list_cache = []

# --- BỘ NÃO XỬ LÝ NGÔN NGỮ ---
def parse_boq_name(boq_name):
    name_lower = str(boq_name).lower().strip()
    dn_val = None
    dn_match = re.search(r'(dn|d|phi|ø|ф)\s*(\d+)', name_lower)
    if dn_match: dn_val = int(dn_match.group(2))

    norm_name = name_lower

    if any(k in name_lower for k in ['van hút', 'van đáy', 'rọ hút']): norm_name = f"rọ hút dn{dn_val}" if dn_val else "rọ hút"
    elif any(k in name_lower for k in ['van cổng', 'van chặn', 'van cửa', 'van xả đáy', 'van cô lập']):
        if 'giám sát' in name_lower or 'tín hiệu' in name_lower: norm_name = f"van cổng os&y kèm công tắc giám sát dn{dn_val}" if dn_val else "van cổng os&y kèm công tắc giám sát"
        elif any(k in name_lower for k in ['os&y', 'osy', 'ty nổi']): norm_name = f"van cổng os&y dn{dn_val}" if dn_val else "van cổng os&y"
        elif any(k in name_lower for k in ['nrs', 'ty chìm']): norm_name = f"van cổng nrs dn{dn_val}" if dn_val else "van cổng nrs"
        else: norm_name = f"van cổng dn{dn_val}" if dn_val and dn_val <= 50 else (f"van cổng nrs dn{dn_val}" if dn_val else "van cổng")
    elif 'van bướm' in name_lower:
        if 'giám sát' in name_lower or 'tín hiệu' in name_lower: norm_name = f"van bướm tín hiệu dn{dn_val}" if dn_val else "van bướm tín hiệu"
        elif 'tay gạt' in name_lower: norm_name = f"van bướm tay gạt dn{dn_val}" if dn_val else "van bướm tay gạt"
        else: norm_name = f"van bướm tay quay dn{dn_val}" if dn_val else "van bướm tay quay"
    elif any(k in name_lower for k in ['1 chiều', 'một chiều', 'lá lật', 'cản ngược']):
        norm_name = f"van một chiều dn{dn_val}" if dn_val else "van một chiều"
        if dn_val: norm_name += " nhỏ" if dn_val <= 50 else " lớn"
    elif any(k in name_lower for k in ['báo động', 'alarm']): norm_name = f"van báo động dn{dn_val}" if dn_val else "van báo động"
    elif any(k in name_lower for k in ['xả tràn', 'deluge']): norm_name = f"van xả tràn dn{dn_val}" if dn_val else "van xả tràn"
    elif any(k in name_lower for k in ['xả khí', 'air valve']): norm_name = f"van xả khí dn{dn_val}" if dn_val else "van xả khí"
    elif any(k in name_lower for k in ['giảm áp', 'prv']): norm_name = f"van giảm áp dn{dn_val}" if dn_val else "van giảm áp"
    elif any(k in name_lower for k in ['xả áp', 'an toàn', 'relief']): norm_name = f"van an toàn dn{dn_val}" if dn_val else "van an toàn"
    elif 'van phao' in name_lower: norm_name = f"van phao dn{dn_val}" if dn_val else "van phao"
    elif any(k in name_lower for k in ['y lọc', 'lọc y', 'strainer']): norm_name = f"y lọc dn{dn_val}" if dn_val else "y lọc"
    elif any(k in name_lower for k in ['nối mềm', 'khớp nối mềm']): norm_name = f"khớp nối mềm dn{dn_val}" if dn_val else "khớp nối mềm"
    elif any(k in name_lower for k in ['đồng hồ áp', 'áp kế']): norm_name = "đồng hồ áp suất"
    elif any(k in name_lower for k in ['công tắc áp', 'rơ le áp']): norm_name = "công tắc áp suất"
    elif any(k in name_lower for k in ['dòng chảy', 'flow switch']): norm_name = "công tắc dòng chảy"
    elif 'sprinkler' in name_lower or 'đầu phun' in name_lower:
        norm_name = "đầu phun"
        if any(k in name_lower for k in ['hướng lên', 'quay lên', 'upright']): norm_name += " hướng lên"
        elif any(k in name_lower for k in ['hướng xuống', 'quay xuống', 'pendent', 'xuống lên']): norm_name += " hướng xuống"
        elif any(k in name_lower for k in ['gắn tường', 'sidewall']): norm_name += " gắn tường"
        
        if any(k in name_lower for k in ['nhanh', 'quick']): norm_name += " phản ứng nhanh"
        else: norm_name += " phản ứng tiêu chuẩn"
            
        t_match = re.search(r'(\d+)\s*(độ|c|°c)', name_lower)
        if t_match: norm_name += f" {t_match.group(1)} độ"
        
        k_match = re.search(r'k\s*=?\s*(\d+\.?\d*)', name_lower)
        if k_match: norm_name += f" k{k_match.group(1)}"
        
    elif any(k in name_lower for k in ['đầu báo khói', 'báo khói', 'smoke detector']): 
        norm_name = "đầu báo khói"
        if 'địa chỉ' in name_lower: norm_name += " địa chỉ"
    elif any(k in name_lower for k in ['đầu báo nhiệt', 'báo nhiệt', 'heat detector']): 
        norm_name = "đầu báo nhiệt"
        if 'địa chỉ' in name_lower: norm_name += " địa chỉ"
        if 'gia tăng' in name_lower: norm_name += " gia tăng"
        if 'cố định' in name_lower: norm_name += " cố định"
    elif any(k in name_lower for k in ['nút nhấn', 'nút ấn', 'call point']): 
        norm_name = "nút nhấn khẩn"
        if 'địa chỉ' in name_lower: norm_name += " địa chỉ"
    elif any(k in name_lower for k in ['chuông', 'fire bell']): norm_name = "chuông báo cháy"
    elif any(k in name_lower for k in ['còi', 'đèn chớp', 'horn strobe']): norm_name = "còi đèn báo cháy"
    elif 'module' in name_lower or 'mô đun' in name_lower:
        if 'giám sát' in name_lower: norm_name = "module giám sát"
        elif 'điều khiển' in name_lower: norm_name = "module điều khiển"
        elif 'cách ly' in name_lower: norm_name = "module cách ly"
        else: norm_name = "module báo cháy"
    elif any(k in name_lower for k in ['tủ trung tâm', 'tủ báo cháy', 'facp']): norm_name = "tủ trung tâm báo cháy"
    elif 'cáp' in name_lower or 'cable' in name_lower or 'dây điện' in name_lower:
        if 'chống cháy' in name_lower or 'fr' in name_lower: norm_name = "cáp chống cháy"
        elif 'chậm cháy' in name_lower: norm_name = "cáp chậm cháy"
        elif 'tín hiệu' in name_lower: norm_name = "cáp tín hiệu"
        else: norm_name = "cáp điện"
        c_match = re.search(r'(\d+)\s*[cx*x]\s*(\d+\.?\d*)', name_lower)
        if c_match: norm_name += f" {c_match.group(1)}x{c_match.group(2)}"
    elif any(k in name_lower for k in ['ty ren', 'ty treo', 'thanh ren']): norm_name = "ty ren"
    elif any(k in name_lower for k in ['cùm', 'đai treo', 'quang treo']): norm_name = f"cùm treo dn{dn_val}" if dn_val else "cùm treo"
    elif any(k in name_lower for k in ['ubolt', 'u-bolt', 'u bolt']): norm_name = f"u-bolt dn{dn_val}" if dn_val else "u-bolt"
        
    return norm_name, dn_val

def fuzzy_match(parsed_name, dn_val, price_dict):
    parsed_name = parsed_name.replace("nhỏ", "").replace("lớn", "").strip()
    if parsed_name in price_dict:
        return price_dict[parsed_name]
        
    keywords = parsed_name.split()
    best_match = None
    max_score = 0
    
    p_k_match = re.search(r'k\s*=?\s*(\d+\.?\d*)', parsed_name)
    p_t_match = re.search(r'(\d+)\s*độ', parsed_name)
    
    for db_name, data in price_dict.items():
        score = sum(1 for kw in keywords if kw in db_name)
        if dn_val:
            db_dn_match = re.search(r'(dn|d|phi|ø)\s*(\d+)', db_name)
            if db_dn_match and int(db_dn_match.group(2)) != dn_val: continue 
        if p_k_match:
            db_k = re.search(r'k\s*=?\s*(\d+\.?\d*)', db_name)
            if db_k and db_k.group(1) != p_k_match.group(1): continue
        if p_t_match:
            db_t = re.search(r'(\d+)\s*(độ|c|°c)', db_name)
            if db_t and db_t.group(1) != p_t_match.group(1): continue
        if score > max_score and score >= len(keywords) * 0.5: 
            max_score = score
            best_match = data
            
    return best_match

# --- MAIN EXECUTION ---
if uploaded_boq is not None:
    try:
        if uploaded_boq.name.endswith('.csv'): df_raw = pd.read_csv(uploaded_boq)
        else: df_raw = pd.read_excel(uploaded_boq, sheet_name=boq_sheet_selected).dropna(how='all').dropna(axis=1, how='all')
        
        if len(df_raw.columns) == 0:
            st.error(f"❌ Sheet '{boq_sheet_selected}' không có dữ liệu.")
            st.stop()

        boq_name_col, boq_qty_col = df_raw.columns[0], df_raw.columns[0]
        for col in df_raw.columns:
            col_lower = str(col).lower()
            if any(k in col_lower for k in ['tên', 'thiết bị', 'mô tả', 'diễn giải', 'tiêu chuẩn']): boq_name_col = col
            if any(k in col_lower for k in ['số lượng', 'sl', 'qty']): boq_qty_col = col

        if st.sidebar.button("🚀 BẮT ĐẦU XỬ LÝ BOQ", type="primary", use_container_width=True):
            with st.spinner("Phần mềm đang phân tích dữ liệu..."):
                columns_template = ['Mã hàng', 'Diễn giải', 'Xuất xứ', 'ĐVT', 'SL', 'Đơn giá (VND)', 'Thành tiền (VND)', 'Ghi chú']
                df_processed = pd.DataFrame(columns=columns_template)
    
                price_dict = {}
                name_col_price = None
                extracted_price_list = [] 
    
                if os.path.exists(PRICE_LIST_PATH) and price_sheet_selected:
                    df_price = pd.read_excel(PRICE_LIST_PATH, sheet_name=price_sheet_selected).dropna(how='all').dropna(axis=1, how='all')
                    if len(df_price.columns) > 0:
                        name_col_price = df_price.columns[0]
                        for col in df_price.columns:
                            if any(k in str(col).lower() for k in ['tên', 'mô tả', 'thiết bị', 'tiêu chuẩn']):
                                name_col_price = col
                                break
                        for _, pr in df_price.iterrows():
                            k = str(pr.get(name_col_price, '')).strip().lower()
                            if k and k != 'nan':
                                price_dict[k] = pr
                                ma = ""
                                for m_col in ['Mã SP', 'Mã hàng', 'Mã sản phẩm']:
                                    if m_col in pr and pd.notna(pr[m_col]): ma = str(pr[m_col]); break
                                dg = 0
                                for p_col in ['Đơn giá', 'Giá', 'Price', 'Đơn giá (VND)']:
                                    if p_col in pr and pd.notna(pr[p_col]): dg = pd.to_numeric(pr[p_col], errors='coerce') or 0; break
                                xx = ""
                                for o_col in ['Xuất xứ', 'XUẤT XỨ', 'Thương hiệu']:
                                    if o_col in pr and pd.notna(pr[o_col]): xx = str(pr[o_col]); break
                                dvt = "Cái"
                                for dvt_col in ['ĐVT', 'Đơn vị tính']:
                                    if dvt_col in pr and pd.notna(pr[dvt_col]): dvt = str(pr[dvt_col]); break
                                
                                extracted_price_list.append({
                                    'Tên': str(pr.get(name_col_price, '')),
                                    'Mã hàng': ma,
                                    'Đơn giá': dg,
                                    'Xuất xứ': xx,
                                    'ĐVT': dvt
                                })
                
                st.session_state.price_list_cache = extracted_price_list
    
                for idx, row in df_raw.iterrows():
                    boq_name = str(row.get(boq_name_col, '')).strip()
                    boq_qty = row.get(boq_qty_col, 0)
                    if not boq_name or boq_name.lower() == 'nan': continue
    
                    new_row = {
                        'Mã hàng': '',
                        'Diễn giải': boq_name, 
                        'Xuất xứ': '',
                        'ĐVT': 'Cái' if 'cáp' not in boq_name.lower() else 'Mét', 
                        'SL': pd.to_numeric(boq_qty, errors='coerce') or 0,
                        'Đơn giá (VND)': 0,
                        'Thành tiền (VND)': 0,
                        'Ghi chú': ''
                    }
    
                    parsed_name, dn_val = parse_boq_name(boq_name)
                    matched = fuzzy_match(parsed_name, dn_val, price_dict)
                    
                    if matched is not None:
                        if name_col_price and pd.notna(matched[name_col_price]):
                            new_row['Diễn giải'] = str(matched[name_col_price])
                        for m_col in ['Mã SP', 'Mã hàng', 'Mã sản phẩm']:
                            if m_col in matched and pd.notna(matched[m_col]): new_row['Mã hàng'] = matched[m_col]; break
                        for p_col in ['Đơn giá', 'Giá', 'Price', 'Đơn giá (VND)']:
                            if p_col in matched and pd.notna(matched[p_col]): new_row['Đơn giá (VND)'] = pd.to_numeric(matched[p_col], errors='coerce') or 0; break
                        for o_col in ['Xuất xứ', 'XUẤT XỨ', 'Thương hiệu']:
                            if o_col in matched and pd.notna(matched[o_col]): new_row['Xuất xứ'] = matched[o_col]; break
                        for dvt_col in ['ĐVT', 'Đơn vị tính']:
                            if dvt_col in matched and pd.notna(matched[dvt_col]): new_row['ĐVT'] = matched[dvt_col]; break
    
                    new_row['Thành tiền (VND)'] = new_row['SL'] * new_row['Đơn giá (VND)']
                    df_processed = pd.concat([df_processed, pd.DataFrame([new_row])], ignore_index=True)
    
                st.session_state.df_editable = df_processed

    except Exception as e:
        st.error(f"Lỗi xử lý file BOQ: {e}")

else:
    st.info("👈 Vui lòng tải lên file Khối Lượng (BOQ) ở menu bên trái để bắt đầu lập báo giá.")

# --- GIAO DIỆN CHỈNH SỬA & XUẤT EXCEL TỐI ƯU ---
if st.session_state.df_editable is not None:
    st.markdown("<hr style='border: 1px solid #4A90E2;'>", unsafe_allow_html=True)
    st.markdown("<h3>🛠 CHI TIẾT KHỐI LƯỢNG & ĐƠN GIÁ</h3>", unsafe_allow_html=True)
    
    # --- CÔNG CỤ TRA CỨU THỦ CÔNG ---
    if len(st.session_state.price_list_cache) > 0:
        with st.expander("🔍 TÌM KIẾM VÀ CHÈN MÃ THỦ CÔNG (Dành cho thiết bị chưa nhận diện được)"):
            st.markdown("<p style='color: #A0A0A0; font-size: 14px;'>Gõ tên hoặc mã vật tư vào ô <b>Tìm kiếm</b>. Sau đó chọn dòng muốn sửa và bấm <b>Áp dụng</b>.</p>", unsafe_allow_html=True)
            
            c1, c2 = st.columns([1, 2])
            with c1:
                row_options = st.session_state.df_editable.index
                selected_row = st.selectbox("1. Chọn dòng thiết bị cần sửa:", options=row_options, 
                                            format_func=lambda x: f"Dòng {x+1}: {st.session_state.df_editable.at[x, 'Diễn giải']}")
            with c2:
                price_opts = {f"[{p['Mã hàng']}] {p['Tên']} - {p['Đơn giá']:,.0f} đ": p for p in st.session_state.price_list_cache}
                selected_item_str = st.selectbox("2. Tìm và Chọn thiết bị trong Bảng giá:", options=["-- Trống --"] + list(price_opts.keys()))
                
            if st.button("✅ ÁP DỤNG MÃ VÀO BẢNG"):
                if selected_item_str != "-- Trống --":
                    p_data = price_opts[selected_item_str]
                    st.session_state.df_editable.at[selected_row, 'Mã hàng'] = p_data['Mã hàng']
                    st.session_state.df_editable.at[selected_row, 'Diễn giải'] = p_data['Tên']
                    st.session_state.df_editable.at[selected_row, 'Xuất xứ'] = p_data['Xuất xứ']
                    st.session_state.df_editable.at[selected_row, 'ĐVT'] = p_data['ĐVT']
                    st.session_state.df_editable.at[selected_row, 'Đơn giá (VND)'] = p_data['Đơn giá']
                    st.session_state.df_editable.at[selected_row, 'Thành tiền (VND)'] = st.session_state.df_editable.at[selected_row, 'SL'] * p_data['Đơn giá']
                    st.rerun()

    edited_df = st.data_editor(
        st.session_state.df_editable, 
        num_rows="dynamic", 
        use_container_width=True,
        height=500,
        column_config={
            "SL": st.column_config.NumberColumn("SL", format="%d"),
            "Đơn giá (VND)": st.column_config.NumberColumn("Đơn giá (VND)", format="%,d"),
            "Thành tiền (VND)": st.column_config.NumberColumn("Thành tiền (VND)", format="%,d", disabled=True)
        }
    )
    st.session_state.df_editable = edited_df

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_out1, col_out2 = st.columns([3, 1])
    with col_out1:
        output_filename = st.text_input("Tên file Excel xuất ra:", value=f"BaoGia_QNTech_{datetime.now().strftime('%Y%m%d')}.xlsx")
    with col_out2:
        st.write("") 
        st.write("")
        if st.button("📥 XUẤT FILE BÁO GIÁ", use_container_width=True):
            try:
                with st.spinner("Đang xử lý thu phóng dòng tự động, bảo vệ cấu trúc Excel..."):
                    wb = openpyxl.load_workbook(TEMPLATE_PATH)
                    ws = wb.active

                    def safe_write(row, col, val):
                        try: ws.cell(row=row, column=col).value = val
                        except Exception: pass

                    start_row = 14
                    num_items = len(edited_df)
                    
                    end_template_row = start_row
                    for r in range(start_row, start_row + 150):
                        val_c2 = str(ws.cell(row=r, column=2).value or "").lower()
                        val_c3 = str(ws.cell(row=r, column=3).value or "").lower()
                        if "cộng" in val_c2 or "tổng" in val_c2 or "cộng" in val_c3 or "tổng" in val_c3:
                            end_template_row = r
                            break
                    
                    available_rows = end_template_row - start_row
                    
                    # === BẢN VÁ LỖI XÓA DÒNG BỊ MẤT FORMAT ===
                    if num_items > available_rows:
                        rows_to_insert = num_items - available_rows
                        for m in list(ws.merged_cells.ranges):
                            if m.min_row >= end_template_row: m.shift(row_shift=rows_to_insert)

                        max_dim = max(list(ws.row_dimensions.keys()) + [0])
                        old_heights = {}
                        for r in range(end_template_row, max_dim + 1):
                            if r in ws.row_dimensions: old_heights[r] = ws.row_dimensions[r].height
                                
                        ws.insert_rows(end_template_row, amount=rows_to_insert)
                        
                        for r, h in old_heights.items(): ws.row_dimensions[r + rows_to_insert].height = h
                            
                        source_row = end_template_row - 1
                        item_height = ws.row_dimensions[source_row].height if source_row in ws.row_dimensions else None
                        
                        for r in range(end_template_row, end_template_row + rows_to_insert):
                            if item_height: ws.row_dimensions[r].height = item_height
                            for col in range(1, 10):
                                s_cell = ws.cell(row=source_row, column=col)
                                t_cell = ws.cell(row=r, column=col)
                                if s_cell.has_style:
                                    t_cell.font = copy.copy(s_cell.font); t_cell.border = copy.copy(s_cell.border)
                                    t_cell.fill = copy.copy(s_cell.fill); t_cell.number_format = copy.copy(s_cell.number_format)
                                    t_cell.alignment = copy.copy(s_cell.alignment)

                        end_template_row += rows_to_insert

                    elif num_items < available_rows:
                        rows_to_delete = available_rows - num_items
                        delete_start = start_row + num_items
                        
                        # 1. Dịch chuyển ô gộp (Merged Cells)
                        new_merged = []
                        for m in list(ws.merged_cells.ranges):
                            if m.min_row > delete_start:
                                m.shift(row_shift=-rows_to_delete)
                            new_merged.append(m)
                            
                        # 2. Sao lưu chiều cao dòng (Row dimensions) của Footer
                        max_dim = max(list(ws.row_dimensions.keys()) + [0])
                        old_heights = {}
                        for r in range(delete_start + rows_to_delete, max_dim + 1):
                            if r in ws.row_dimensions:
                                old_heights[r] = ws.row_dimensions[r].height
                                
                        # 3. Tiến hành Xóa dòng
                        ws.delete_rows(delete_start, amount=rows_to_delete)
                        
                        # 4. Gắn lại chiều cao cho Footer
                        for r in range(delete_start, max_dim - rows_to_delete + 1):
                            ws.row_dimensions[r].height = None  # Xóa cache chiều cao cũ
                        for r, h in old_heights.items():
                            ws.row_dimensions[r - rows_to_delete].height = h
                            
                        end_template_row -= rows_to_delete

                    # --- Ghi dữ liệu ---
                    for idx, row in edited_df.iterrows():
                        current_row = start_row + idx
                        safe_write(current_row, 1, idx + 1)
                        safe_write(current_row, 2, row.get('Mã hàng', ''))
                        safe_write(current_row, 3, row.get('Diễn giải', ''))
                        safe_write(current_row, 4, row.get('Xuất xứ', ''))
                        safe_write(current_row, 5, row.get('ĐVT', 'Cái'))
                        safe_write(current_row, 6, row.get('SL', 0))
                        safe_write(current_row, 7, row.get('Đơn giá (VND)', 0))
                        try: ws.cell(row=current_row, column=8).value = f"=F{current_row}*G{current_row}"
                        except Exception: pass
                        safe_write(current_row, 9, row.get('Ghi chú', ''))

                    try: ws.cell(row=end_template_row, column=8).value = f"=SUM(H{start_row}:H{end_template_row-1})"
                    except Exception: pass

                    output_buffer = BytesIO()
                    wb.save(output_buffer)
                    output_buffer.seek(0)

                    st.success("🎉 GHI DỮ LIỆU THÀNH CÔNG! Bấm tải xuống bên dưới.")
                    st.download_button(
                        label="⬇️ NHẤN ĐỂ LƯU FILE EXCEL",
                        data=output_buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Đã xảy ra lỗi hệ thống: {e}")