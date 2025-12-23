import streamlit as st
import pandas as pd
import io
import datetime

# --- CẤU HÌNH TRANG ---
st.set_page_config(page_title="Tool Hồ Sơ Điện Lực (Full Sign)", layout="wide", page_icon="🖨️")

# --- KHỞI TẠO BỘ NHỚ ---
if 'projects' not in st.session_state: st.session_state.projects = [] 
if 'current_items' not in st.session_state: st.session_state.current_items = []

# --- HÀM CHUYỂN SỐ SANG LA MÃ ---
def to_roman(n):
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
    roman_num = ''
    i = 0
    while  n > 0:
        for _ in range(n // val[i]):
            roman_num += syb[i]
            n -= val[i]
        i += 1
    return roman_num

# --- HÀM XỬ LÝ SỐ LIỆU ---
def clean_num(x):
    try:
        if pd.isna(x): return 0.0
        return float(str(x).replace(',', '').replace('.', '').strip())
    except: return 0.0

@st.cache_data
def load_price_list_advanced(file):
    try:
        if file.name.endswith('.csv'): df = pd.read_csv(file, header=9)
        else: df = pd.read_excel(file, header=9)
        
        items = []
        for index, row in df.iterrows():
            ma_vt = str(row.iloc[2]).strip()
            ten_vt = str(row.iloc[4]).strip()
            dvt = str(row.iloc[5]).strip()
            
            if ten_vt == 'nan' or ten_vt == '': continue
            
            # TỒN KHO
            sl_ton = clean_num(row.iloc[7])
            gia_ton = clean_num(row.iloc[8])
            if sl_ton > 0:
                label = f"📦 [TỒN KHO] {ten_vt} (SL: {sl_ton:,.0f}) - Giá: {gia_ton:,.0f}"
                items.append([ma_vt, ten_vt, dvt, "Tồn kho", gia_ton, label])
            
            # HỢP ĐỒNG
            gia_hd = clean_num(row.iloc[11])
            if gia_hd > 0:
                label = f"📝 [HỢP ĐỒNG] {ten_vt} - Giá: {gia_hd:,.0f}"
                items.append([ma_vt, ten_vt, dvt, "Hợp đồng", gia_hd, label])
                
            # BÁN LẺ
            gia_le = clean_num(row.iloc[17])
            if gia_le > 0:
                label = f"💰 [BÁN LẺ] {ten_vt} - Giá: {gia_le:,.0f}"
                items.append([ma_vt, ten_vt, dvt, "Bán lẻ", gia_le, label])

        return pd.DataFrame(items, columns=["Mã VT", "Tên Gốc", "ĐVT", "Loại Giá", "Đơn Giá", "Hiển Thị"])
    except: return None

# --- GIAO DIỆN CHÍNH ---
st.title("🖨️ CÔNG CỤ TẠO HỒ SƠ ĐIỆN LỰC (FORM CHUẨN)")
st.caption("Xuất file Excel với định dạng văn bản hành chính & Đầy đủ chữ ký")
st.markdown("---")

col_left, col_right = st.columns([1, 2])

# --- CỘT TRÁI: NHẬP LIỆU ---
with col_left:
    st.header("1. Nhập Liệu")
    price_file = st.file_uploader("File Giá (.xlsx, .csv)", type=['csv', 'xlsx'])
    
    if st.button("🗑️ Xóa làm lại", type="primary"):
        st.session_state.projects = []
        st.session_state.current_items = []
        st.rerun()
    
    st.divider()

    if price_file:
        df_pro = load_price_list_advanced(price_file)
        
        if df_pro is not None:
            prj_name = st.text_input("Tên Trạm / Hạng mục:", placeholder="VD: Trạm T1 Phước Đông")
            
            selected_label = st.selectbox("Chọn vật tư:", options=df_pro["Hiển Thị"], index=None)
            
            c1, c2, c3 = st.columns(3)
            qty_new = c1.number_input("Thay Mới", min_value=0.0, step=1.0)
            qty_reuse = c2.number_input("Tận Dụng", min_value=0.0, step=1.0)
            qty_rec = c3.number_input("Thu Hồi", min_value=0.0, step=1.0)
            note = st.text_input("Ghi chú:")
            
            if st.button("➕ Thêm"):
                if selected_label:
                    item_data = df_pro[df_pro["Hiển Thị"] == selected_label].iloc[0]
                    st.session_state.current_items.append({
                        "Mã VT": item_data["Mã VT"],
                        "Tên VTTB": item_data["Tên Gốc"],
                        "ĐVT": item_data["ĐVT"],
                        "Nguồn Giá": item_data["Loại Giá"],
                        "Đơn Giá": item_data["Đơn Giá"],
                        "Thay Mới": qty_new,
                        "Tận Dụng": qty_reuse,
                        "Thu Hồi": qty_rec,
                        "Ghi Chú": note
                    })
                    st.toast(f"Đã thêm: {item_data['Tên Gốc']}")
    else:
        st.info("👈 Nạp file Giá trước.")

# --- CỘT PHẢI: XEM & XUẤT ---
with col_right:
    st.header("2. Xem & Xuất Hồ Sơ")
    
    if price_file:
        # Bảng xem trước
        if st.session_state.current_items:
            st.caption(f"Đang nhập: **{prj_name}**")
            df_curr = pd.DataFrame(st.session_state.current_items)
            st.dataframe(df_curr[["Tên VTTB", "Thay Mới", "Tận Dụng", "Thu Hồi"]], use_container_width=True, hide_index=True)
            
            if st.button("💾 LƯU TRẠM NÀY"):
                if prj_name:
                    st.session_state.projects.append({"name": prj_name, "data": df_curr})
                    st.session_state.current_items = []
                    st.rerun()
                else: st.warning("Nhập tên trạm!")

        st.divider()
        if st.session_state.projects:
            st.success(f"✅ Đã lưu {len(st.session_state.projects)} trạm.")
            
            with st.expander("⚙️ CẤU HÌNH VĂN BẢN & CHỮ KÝ", expanded=True):
                col_h1, col_h2 = st.columns(2)
                with col_h1:
                    so_phuong_an = st.text_input("Số Phương án:", value="....../PA-PCTN")
                    ngay_thang = st.date_input("Ngày lập:", datetime.date.today())
                    dia_diem = st.text_input("Địa điểm:", value="Tây Ninh")
                
                with col_h2:
                    nguoi_lap = st.text_input("Người lập:", value="Nguyễn Văn A")
                    to_kt = st.text_input("Tổ Kỹ Thuật:", value="Trần Văn B")
                    lanh_dao = st.text_input("Giám Đốc:", value="Ông Lãnh Đạo")

            if st.button("📥 XUẤT FILE EXCEL (CHUẨN FORM)", type="primary"):
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                wb = writer.book
                
                # --- ĐỊNH DẠNG STYLE ---
                s_base = {'font_name': 'Times New Roman', 'font_size': 13}
                
                # Styles
                f_header_left = wb.add_format({**s_base, 'bold': False, 'align': 'center', 'valign': 'top', 'text_wrap': True})
                f_header_right = wb.add_format({**s_base, 'bold': True, 'align': 'center', 'valign': 'top', 'text_wrap': True})
                f_date = wb.add_format({**s_base, 'italic': True, 'align': 'center'})
                f_title = wb.add_format({**s_base, 'bold': True, 'font_size': 14, 'align': 'center'})
                f_subtitle = wb.add_format({**s_base, 'italic': True, 'align': 'center'})
                f_th = wb.add_format({**s_base, 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                f_td_center = wb.add_format({**s_base, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                f_td_roman_bold = wb.add_format({**s_base, 'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                f_td_left = wb.add_format({**s_base, 'border': 1, 'align': 'left', 'valign': 'vcenter', 'indent': 1, 'text_wrap': True})
                f_item_name = wb.add_format({**s_base, 'bold': True, 'border': 1, 'align': 'left', 'valign': 'vcenter', 'indent': 1})
                f_money = wb.add_format({**s_base, 'border': 1, 'num_format': '#,##0', 'valign': 'vcenter'})
                f_sign_title = wb.add_format({**s_base, 'bold': True, 'align': 'center'})
                f_sign_name = wb.add_format({**s_base, 'bold': True, 'align': 'center'})

                all_summary = {} 
                
                # ==========================
                # SHEET 1: BẢNG KÊ VTTB
                # ==========================
                ws = wb.add_worksheet("BANG_KE_VTTB")
                ws.set_paper(9) # A4
                ws.set_margins(0.7, 0.7, 0.75, 0.75)
                
                # Header Page 1
                ws.merge_range("A1:C1", "TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM\nCÔNG TY ĐIỆN LỰC TÂY NINH\n-------", f_header_left)
                ws.merge_range("D1:G1", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n---------------", f_header_right)
                ws.set_row(0, 60) 

                ws.merge_range("A2:C2", f"Số: {so_phuong_an}", f_header_left)
                ws.merge_range("D2:G2", f"{dia_diem}, ngày {ngay_thang.day} tháng {ngay_thang.month} năm {ngay_thang.year}", f_date)
                
                curr = 4
                ws.merge_range(curr, 0, curr, 6, "BẢNG LIỆT KÊ VẬT TƯ THIẾT BỊ", f_title)
                curr += 1
                ws.merge_range(curr, 0, curr, 6, f"(Kèm theo P.án số: {so_phuong_an})", f_subtitle)
                curr += 2
                
                headers1 = ["Stt", "Tên vật tư - Thiết bị", "ĐVT", "Thay mới", "Tận dụng", "Thu hồi", "Ghi chú"]
                for c, h in enumerate(headers1): ws.write(curr, c, h, f_th)
                curr += 1
                
                # Loop Data Sheet 1
                has_items_b1 = False
                for i, p in enumerate(st.session_state.projects):
                    df = p['data']
                    df_vttb = df[(df["Thay Mới"] > 0) | (df["Tận Dụng"] > 0)].copy()
                    
                    if not df_vttb.empty:
                        has_items_b1 = True
                        roman = to_roman(i+1)
                        ws.write(curr, 0, roman, f_td_roman_bold)
                        ws.merge_range(curr, 1, curr, 6, p['name'], f_item_name)
                        curr += 1
                        
                        for idx, row in df_vttb.reset_index(drop=True).iterrows():
                            ws.write(curr, 0, idx+1, f_td_center)
                            ws.write(curr, 1, row['Tên VTTB'], f_td_left)
                            ws.write(curr, 2, row['ĐVT'], f_td_center)
                            ws.write(curr, 3, row['Thay Mới'] if row['Thay Mới'] > 0 else "", f_td_center)
                            ws.write(curr, 4, row['Tận Dụng'] if row['Tận Dụng'] > 0 else "", f_td_center)
                            ws.write(curr, 5, "", f_td_center)
                            ws.write(curr, 6, row['Ghi Chú'], f_td_center)
                            curr += 1

                if not has_items_b1:
                    ws.merge_range(curr, 0, curr, 6, "(Không có)", f_td_center)
                    curr += 1

                # Loop Data Thu Hồi
                curr += 2
                ws.merge_range(curr, 0, curr, 6, "BẢNG LIỆT KÊ VẬT TƯ THU HỒI", f_title)
                curr += 1
                
                ws.write(curr, 0, "Stt", f_th)
                ws.write(curr, 1, "Tên vật tư thu hồi", f_th)
                ws.write(curr, 2, "ĐVT", f_th)
                ws.merge_range(curr, 3, curr, 5, "Số lượng", f_th)
                ws.write(curr, 6, "Ghi chú", f_th)
                curr += 1
                
                has_items_b2 = False
                for i, p in enumerate(st.session_state.projects):
                    df = p['data']
                    df_thuhoi = df[df["Thu Hồi"] > 0].copy()
                    
                    if not df_thuhoi.empty:
                        has_items_b2 = True
                        roman = to_roman(i+1)
                        ws.write(curr, 0, roman, f_td_roman_bold)
                        ws.merge_range(curr, 1, curr, 6, p['name'], f_item_name)
                        curr += 1
                        
                        for idx, row in df_thuhoi.reset_index(drop=True).iterrows():
                            ws.write(curr, 0, idx+1, f_td_center)
                            ws.write(curr, 1, row['Tên VTTB'], f_td_left)
                            ws.write(curr, 2, row['ĐVT'], f_td_center)
                            ws.merge_range(curr, 3, curr, 5, row['Thu Hồi'], f_td_center)
                            ws.write(curr, 6, row['Ghi Chú'], f_td_center)
                            curr += 1
                            
                if not has_items_b2:
                    ws.merge_range(curr, 0, curr, 6, "(Không có vật tư thu hồi)", f_td_center)
                    curr += 1

                # Chữ ký Sheet 1
                curr += 3
                ws.write(curr, 1, "LẬP BẢNG", f_sign_title)
                ws.write(curr, 3, "PHÒNG KỸ THUẬT", f_sign_title)
                ws.merge_range(curr, 4, curr, 6, "GIÁM ĐỐC", f_sign_title)
                
                curr += 5
                ws.write(curr, 1, nguoi_lap, f_sign_name)
                ws.write(curr, 3, to_kt, f_sign_name)
                ws.merge_range(curr, 4, curr, 6, lanh_dao, f_sign_name)

                # Set Widths Sheet 1
                ws.set_column(0, 0, 6)
                ws.set_column(1, 1, 40)
                ws.set_column(2, 6, 12)

                # ==========================
                # SHEET 2: TỔNG HỢP CHUNG
                # ==========================
                for p in st.session_state.projects:
                    for _, r in p['data'].iterrows():
                        if r["Thay Mới"] > 0:
                            key = (r["Mã VT"], r["Tên VTTB"], r["ĐVT"], r["Nguồn Giá"], r["Đơn Giá"])
                            if key in all_summary: all_summary[key] += r["Thay Mới"]
                            else: all_summary[key] = r["Thay Mới"]

                ws_sum = wb.add_worksheet("TONG_HOP_CHUNG")
                ws_sum.set_paper(9)
                ws_sum.set_margins(0.7, 0.7, 0.75, 0.75)

                # Header Sheet 2
                ws_sum.merge_range("A1:C1", "TỔNG CÔNG TY ĐIỆN LỰC MIỀN NAM\nCÔNG TY ĐIỆN LỰC TÂY NINH\n-------", f_header_left)
                ws_sum.merge_range("D1:H1", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n---------------", f_header_right)
                ws_sum.set_row(0, 60)
                
                ws_sum.merge_range(4, 0, 4, 7, "BẢNG TỔNG HỢP KHỐI LƯỢNG VÀ CHIẾT TÍNH", f_title)
                
                hs = ["STT", "Mã VT", "Tên Vật Tư", "Nguồn Giá", "ĐVT", "Số Lượng", "Đơn Giá", "Thành Tiền"]
                for c, t in enumerate(hs): ws_sum.write(6, c, t, f_th)
                
                ridx = 7
                stt = 1
                total = 0
                for (ma, ten, dvt, nguon, gia), sl in sorted(all_summary.items(), key=lambda x: x[0][1]):
                    tt = sl * gia
                    total += tt
                    ws_sum.write(ridx, 0, stt, f_td_center)
                    ws_sum.write(ridx, 1, ma, f_td_center)
                    ws_sum.write(ridx, 2, ten, f_td_left)
                    ws_sum.write(ridx, 3, nguon, f_td_center)
                    ws_sum.write(ridx, 4, dvt, f_td_center)
                    ws_sum.write(ridx, 5, sl, f_td_center)
                    ws_sum.write(ridx, 6, gia, f_money)
                    ws_sum.write(ridx, 7, tt, f_money)
                    ridx += 1
                    stt += 1
                
                # Tổng cộng
                ws_sum.merge_range(ridx, 0, ridx, 6, "TỔNG CỘNG (Chưa VAT):", wb.add_format({**s_base, 'bold': True, 'align': 'right', 'border': 1}))
                ws_sum.write(ridx, 7, total, wb.add_format({**s_base, 'bold': True, 'border': 1, 'num_format': '#,##0', 'bg_color': 'yellow'}))
                
                # --- CHỮ KÝ SHEET 2 (MỚI THÊM) ---
                ridx += 3
                ws_sum.write(ridx, 2, "LẬP BẢNG", f_sign_title) # Căn cột Tên vật tư cho đẹp
                ws_sum.write(ridx, 4, "PHÒNG KỸ THUẬT", f_sign_title) # Căn cột ĐVT
                ws_sum.merge_range(ridx, 6, ridx, 7, "GIÁM ĐỐC", f_sign_title) # Cột tiền

                ridx += 5
                ws_sum.write(ridx, 2, nguoi_lap, f_sign_name)
                ws_sum.write(ridx, 4, to_kt, f_sign_name)
                ws_sum.merge_range(ridx, 6, ridx, 7, lanh_dao, f_sign_name)

                # Set Widths Sheet 2
                ws_sum.set_column(0, 1, 10)
                ws_sum.set_column(2, 2, 40)
                ws_sum.set_column(3, 5, 12)
                ws_sum.set_column(6, 7, 18)

                writer.close()
                st.download_button("📥 TẢI FILE EXCEL CHUẨN", output.getvalue(), f"Ho_So_VTTB_{datetime.date.today()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")