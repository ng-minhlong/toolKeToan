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

    required_columns = {'STT', 'Nguồn', 'Mã đại lý', 'Doanh thu thực thu'}
    if not required_columns.issubset(df.columns):
        st.warning(f"Sheet '{sheet_name}' thiếu một trong các cột: {', '.join(required_columns)}")
        return None

    df_filtered = df[df['STT'].apply(is_valid_stt)].copy()


    # Làm sạch rồi chuyển đổi sang số (sử dụng kiểu tiếng Anh với dấu , là ngăn nhóm và . là thập phân)
    # Làm sạch
    df_filtered['Nguồn'] = (
        df_filtered['Nguồn']
        .astype(str)
        .str.strip()
        .replace("-", "", regex=False)
        .replace("", pd.NA)
        .str.replace(",", "", regex=False)  # Xóa dấu phẩy ngăn nhóm
    )

    # Ép kiểu số
    df_filtered['Nguồn'] = pd.to_numeric(df_filtered['Nguồn'], errors='coerce')
    df_filtered['Doanh thu thực thu'] = pd.to_numeric(df_filtered['Doanh thu thực thu'].astype(str).str.replace(',', ''), errors='coerce')

    # Xử lý dữ liệu theo mã đại lý nếu được cung cấp
    dai_ly_data = {}
    if ma_dai_ly_list:
        for ma_dl in ma_dai_ly_list:
            df_dai_ly = df_filtered[df_filtered['Mã đại lý'] == ma_dl]
            if not df_dai_ly.empty:
                nguon = df_dai_ly['Nguồn'].sum()
                thuc_thu = df_dai_ly['Doanh thu thực thu'].sum()
                ti_le = (thuc_thu / nguon * 100) if nguon != 0 else 0
                dai_ly_data[ma_dl] = {
                    'nguon': nguon,
                    'thuc_thu': thuc_thu,
                    'ti_le': ti_le
                }

    total = df_filtered['Nguồn'].sum()
    truong_phong = total * 0.01  # 1% doanh thu
    team_lead = total * 0.005    # 0.5% doanh thu
    
    st.write(f"Tổng 'Nguồn': **{total:,.0f}**")
    st.write(f"Doanh thu chia Trưởng phòng (1%): **{truong_phong:,.0f}**")
    st.write(f"Doanh thu chia Team Lead (0.5%): **{team_lead:,.0f}**")
    st.write("Xem trước dữ liệu cột 'STT' và 'Nguồn':")
    st.dataframe(df_filtered[['STT', 'Nguồn']].head(300))

    return {
        'total': total,
        'truong_phong': truong_phong,
        'team_lead': team_lead,
        'dai_ly_data': dai_ly_data
    }








# Giao diện Streamlit
st.title("So sánh dữ liệu từ nhiều sheet Excel")

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
            df_compare.columns = ["Tổng Nguồn", "Doanh thu Trưởng phòng", "Doanh thu Team Lead"]
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
