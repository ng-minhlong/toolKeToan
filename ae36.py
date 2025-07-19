import streamlit as st
import pandas as pd

def is_valid_stt(value):
    """Hàm kiểm tra STT là số nguyên hợp lệ"""
    try:
        # Cho phép số float như 1.0 vẫn tính là 1
        val = float(value)
        return val.is_integer() and val > 0
    except:
        return False




def process_sheet(df, sheet_name, ma_dai_ly_list=None):
    st.write(f"### Sheet: {sheet_name}")

    required_columns = {'STT', 'Mã đại lý', 'Doanh thu thực thu'}
    if not required_columns.issubset(df.columns):
        st.warning(f"Sheet '{sheet_name}' thiếu một trong các cột: {', '.join(required_columns)}")
        return None

    df_filtered = df[df['STT'].apply(is_valid_stt)].copy()


    # Làm sạch rồi chuyển đổi sang số (sử dụng kiểu tiếng Anh với dấu , là ngăn nhóm và . là thập phân)
    # Làm sạch
    df_filtered['Doanh thu thực thu'] = (
        df_filtered['Doanh thu thực thu']
        .astype(str)
        .str.strip()
        .replace("-", "", regex=False)
        .replace("", pd.NA)
        .str.replace(",", "", regex=False)  # Xóa dấu phẩy ngăn nhóm
    )

    # Ép kiểu số
    df_filtered['Doanh thu thực thu'] = pd.to_numeric(df_filtered['Doanh thu thực thu'], errors='coerce')
    df_filtered['Doanh thu thực thu'] = pd.to_numeric(df_filtered['Doanh thu thực thu'].astype(str).str.replace(',', ''), errors='coerce')

    # Xử lý dữ liệu theo mã đại lý nếu được cung cấp
    dai_ly_data = {}
    if ma_dai_ly_list:
        for ma_dl in ma_dai_ly_list:
            df_dai_ly = df_filtered[df_filtered['Mã đại lý'] == ma_dl]
            if not df_dai_ly.empty:
                tong_sheet = df_filtered['Doanh thu thực thu'].sum()  # Tổng doanh thu của cả sheet
                thuc_thu = df_dai_ly['Doanh thu thực thu'].sum()     # Doanh thu của đại lý
                ti_le = (thuc_thu / tong_sheet * 100) if tong_sheet != 0 else 0
                dai_ly_data[ma_dl] = {
                    'nguon': tong_sheet,  # Lưu tổng doanh thu của sheet
                    'thuc_thu': thuc_thu,
                    'ti_le': ti_le
                }

    total = df_filtered['Doanh thu thực thu'].sum()
    truong_phong = total * 0.01  # 1% doanh thu
    team_lead = total * 0.005    # 0.5% doanh thu
    quy_van_hanh_cttdtdbh = total * 0.005  # Quỹ vận hành các chương trình thúc đẩy thi đua bán hàng, gắn kết = 0.5% doanh thu

    st.write(f"Tổng 'Doanh thu thực thu': **{total:,.0f}**")
    st.write(f"Doanh thu chia Trưởng phòng (1%): **{truong_phong:,.0f}**")
    st.write(f"Doanh thu chia Team Lead (0.5%): **{team_lead:,.0f}**")
    st.write(f"Quỹ vận hành các chương trình thúc đẩy thi đua bán hàng, gắn kết (0.5%): **{quy_van_hanh_cttdtdbh:,.0f}**")
        # Tính tổng của các dòng "Loại hình nghiệp vụ gốc" bắt đầu bằng "XO"
    if 'Loại hình nghiệp vụ gốc' in df.columns:
        df_filtered['Loại hình nghiệp vụ gốc'] = df['Loại hình nghiệp vụ gốc'].astype(str).str.strip()
        xo_rows = df_filtered[df_filtered['Loại hình nghiệp vụ gốc'].str.startswith("XO")]
        tong_xo = xo_rows['Doanh thu thực thu'].sum()
        ti_le_15 = tong_xo * 0.015
        st.write(f"Chi phí đánh giá rủi ro và cấp đơn( bắt đầu bằng 'XO'): **{ti_le_15:,.0f}**")
    else:
        st.warning("Không tìm thấy cột 'Loại hình nghiệp vụ gốc' để tính tổng mã XO.")

        # Phân loại nghiệp vụ gốc XO / không XO và tính theo "Nguồn đơn vị"
    if sheet_name == "tele HN":
        if 'Loại hình nghiệp vụ gốc' in df.columns and 'Nguồn đơn vị' in df.columns and 'Mã đại lý' in df.columns:
            df_temp = df.copy()
            df_temp['Loại hình nghiệp vụ gốc'] = df_temp['Loại hình nghiệp vụ gốc'].astype(str).str.strip()
            df_temp['Nguồn đơn vị'] = pd.to_numeric(
                df_temp['Nguồn đơn vị'].astype(str).str.replace(",", "").str.strip(),
                errors='coerce'
            )
            
            # Gắn nhãn XO hoặc Non-XO
            df_temp['Nhóm XO'] = df_temp['Loại hình nghiệp vụ gốc'].str.startswith("XO")
            
            # Áp dụng hệ số theo nhóm
            df_temp['Giá trị quy đổi'] = df_temp.apply(
                lambda row: row['Nguồn đơn vị'] * (0.965 if row['Nhóm XO'] else 0.98),
                axis=1
            )
            
            # Nhóm theo Mã đại lý và tính tổng
            df_grouped = df_temp.groupby(['Mã đại lý', 'Nhóm XO'])['Giá trị quy đổi'].sum().reset_index()
            
            st.write("### ✅ Tổng giá trị quy đổi theo Mã đại lý:")
            df_pivot = df_grouped.pivot_table(
                index='Mã đại lý',
                columns='Nhóm XO',
                values='Giá trị quy đổi',
                fill_value=0
            ).rename(columns={True: "XO", False: "Không bắt đầu bằng XO"})
            
            df_pivot["Tổng"] = df_pivot["XO"] + df_pivot["Không bắt đầu bằng XO"]
            st.dataframe(df_pivot.style.format("{:,.0f}"))
        else:
            st.warning("Không đủ cột để tính toán nhóm XO / Không XO trong sheet 'tele HN'")


    
    
    st.write("Xem trước dữ liệu cột 'STT' và 'Doanh thu thực thu':")
    st.dataframe(df_filtered[['STT', 'Doanh thu thực thu']].head(300))

    return {
        'total': total,
        'truong_phong': truong_phong,
        'team_lead': team_lead,
        'dai_ly_data': dai_ly_data
    }








# Giao diện Streamlit
st.title("Phân tích dữ liệu ")

# Input cho mã đại lý
ma_dai_ly_input = st.text_input("Nhập các mã đại lý (phân cách bằng dấu phẩy)", "")
ma_dai_ly_list = [ma.strip() for ma in ma_dai_ly_input.split(',')] if ma_dai_ly_input else []

uploaded_file = st.file_uploader("Tải lên file Excel", type=['xls', 'xlsx'])

if uploaded_file:
    try:
        # Đọc tất cả các sheet
        xls = pd.read_excel(uploaded_file, sheet_name=None, engine='xlrd')
        st.success("Đã tải thành công file Excel!")
        
        
        # Danh sách sheet cần xử lý
        target_sheets = ["Telco", "tele HN", "Tele HCM"]
        totals = {}

        all_dai_ly_data = {}
        
        for sheet in target_sheets:
            if sheet in xls:
                results = process_sheet(xls[sheet], sheet, ma_dai_ly_list)
                if results is not None:
                    totals[sheet] = [results['total'], results['truong_phong'], results['team_lead']]
                    # Lưu thông tin đại lý
                    if ma_dai_ly_list:
                        all_dai_ly_data[sheet] = results['dai_ly_data']
            else:
                st.warning(f"Không tìm thấy sheet: {sheet}")

        # So sánh tổng doanh thu
        if len(totals) == len(target_sheets):
            st.write("## 🔍 So sánh các loại doanh thu")
            df_compare = pd.DataFrame.from_dict(totals, orient='index')
            df_compare.columns = ["Tổng Doanh thu", "Doanh thu Trưởng phòng", "Doanh thu Team Lead"]
            st.dataframe(df_compare.style.format("{:,.0f}"))

        # Hiển thị thông tin và biểu đồ cho các đại lý
        if ma_dai_ly_list and ma_dai_ly_list[0]:  # Kiểm tra có mã đại lý được nhập không
            st.write("## 📊 Phân tích theo Mã đại lý")
            
            # Tạo DataFrame cho việc so sánh
            dai_ly_comparison = []
            for ma_dl in ma_dai_ly_list:
                for sheet in target_sheets:
                    if sheet in all_dai_ly_data and ma_dl in all_dai_ly_data[sheet]:
                        data = all_dai_ly_data[sheet][ma_dl]
                        dai_ly_comparison.append({
                            'Mã đại lý': ma_dl,
                            'Sheet': sheet,
                            'Nguồn': data['nguon'],
                            'Thực thu': data['thuc_thu'],
                            'Tỷ lệ (%)': data['ti_le']
                        })
            
            if dai_ly_comparison:
                df_dai_ly = pd.DataFrame(dai_ly_comparison)
                
                # Hiển thị bảng số liệu
                st.write("### Bảng số liệu chi tiết:")
                st.dataframe(df_dai_ly.style.format({
                    'Nguồn': '{:,.0f}',
                    'Thực thu': '{:,.0f}',
                    'Tỷ lệ (%)': '{:.2f}'
                }))
                
                # Vẽ biểu đồ so sánh doanh thu thực thu
                st.write("### Biểu đồ so sánh doanh thu thực thu:")
                chart_data = pd.pivot_table(
                    df_dai_ly,
                    values='Thực thu',
                    index='Sheet',
                    columns='Mã đại lý',
                    fill_value=0
                )
                st.bar_chart(chart_data)

    except Exception as e:
        st.error(f"Lỗi khi xử lý file Excel: {e}")
