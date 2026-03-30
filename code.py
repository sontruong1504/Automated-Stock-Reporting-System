import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Image
from io import BytesIO
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_JUSTIFY
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from datetime import timedelta
import os
import plotly.express as px


# Cấu hình giao diện
st.set_page_config(page_title="NGUYỄN SƠN TRƯỜNG - K224141742", layout="centered")

# Đổi font chữ Streamlit
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');

    html, body, [class*="st-"] {
        font-family: 'Poppins', sans-serif;
        background-color: #121212;
        color: #FFFFFF;
    }

    h1 {
        font-size: 40px !important;
        font-weight: 700 !important;
        color: #FFFFFF;
        text-align: center;
    }

    h2 {
        font-size: 32px !important;
        font-weight: 600 !important;
        color: #E0E0E0;
        text-align: center;
    }

    p, label, div {
        font-size: 18px !important;
        font-weight: 500 !important;
        color: #F5F5F5;
    }

    .stSelectbox, .stTextInput {
        background-color: #1E1E1E;
        color: #FFFFFF;
        font-size: 18px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

#Chọn màu cho bản báo cáo
red_text = colors.HexColor("#D32F2F ")
bg = colors.HexColor("#D3D3D3")
grey = colors.HexColor("#4B4B4B")
# Hàm load dữ liệu
@st.cache_data
def load_data():
    try:
        base_path = os.path.dirname(__file__)  # lấy thư mục chứa code.py
        # In ra thư mục và file trong đó để kiểm tra
        print("📁 Danh sách file trong thư mục:", os.listdir(base_path))

        df_basic = pd.read_excel(os.path.join(base_path, "basic_info.xlsx"))
        df_info = pd.read_excel(os.path.join(base_path, "Info.xlsx"))
        df_price = pd.read_csv(os.path.join(base_path, "Price.csv"), dtype={"Code": str})  # Sử dụng .csv thay vì .xlsx
        df_price.set_index("Code", inplace=True)
        df_price = df_price.T
        df_price.index = pd.to_datetime(df_price.index)
        df_marketcap = pd.read_csv(os.path.join(base_path, "marketcap.csv"))
        df_ratio = pd.read_excel(os.path.join(base_path, "ratio.xlsx"))
        bcdkt_df = pd.read_csv(os.path.join(base_path, "BCDKT.csv"))  
        kqkd_df = pd.read_csv(os.path.join(base_path, "KQKD.csv"))  
        lctt_df = pd.read_csv(os.path.join(base_path, "LCTT.csv"))
        

        return df_basic, df_info, df_price, df_ratio, bcdkt_df, kqkd_df, lctt_df, df_marketcap

    except FileNotFoundError as e:
        print(f"❌ Không tìm thấy file: {e.filename}")
        st.error(f"❌ Không tìm thấy file: {e.filename}")
        return None

# Load dữ liệu với xử lý lỗi
loaded_data = load_data()
if loaded_data:
    df, info_df, price_df, ratio_df, bcdkt_df, kqkd_df, lctt_df , df_marketcap= loaded_data
else:
    st.stop()

# Lấy danh sách ngày hợp lệ từ file Price
min_date, max_date = price_df.index.min(), price_df.index.max()
st.write(f"Ngày hợp lệ từ: {min_date} đến {max_date}")

# Lấy danh sách ngày hợp lệ từ file Price.csv
min_date, max_date = price_df.index.min(), price_df.index.max()

# Lấy thư mục chứa file hiện tại
base_path = os.path.dirname(__file__)

# Đường dẫn tới các font
font_path_black = os.path.join(base_path, "Roboto_Condensed-Black.ttf")
font_path_regular = os.path.join(base_path, "Roboto_SemiCondensed-Regular.ttf")



# Khai báo font
pdfmetrics.registerFont(TTFont("Roboto_Black", font_path_black))
pdfmetrics.registerFont(TTFont("Roboto_Regular", font_path_regular))




def plot_stock_price(stock_code):
    if stock_code not in price_df.columns:
        print(f"Mã {stock_code} không tồn tại trong dữ liệu.")
        return None

    stock_price = price_df[stock_code].dropna()

    if stock_price.empty:
        print(f"Không có dữ liệu đủ để vẽ biểu đồ cho {stock_code}.")
        return None

    # Vẽ biểu đồ
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(stock_price.index, stock_price.values, linestyle='-', color='#D32F2F')  # Màu đỏ
    ax.grid(True)

    # Định dạng trục y với dấu phẩy
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))

    # Hiển thị năm trên trục X
    years = stock_price.index.year.unique()
    ax.set_xticks([pd.Timestamp(year=year, month=1, day=1) for year in years])
    ax.set_xticklabels(years)
    ax.set_aspect('auto')

    # Chỉnh trục y sang bên phải
    ax.yaxis.set_label_position('right')  # Đổi vị trí nhãn trục y sang bên phải
    ax.yaxis.tick_right()  # Đưa các dấu tick của trục y sang bên phải

    # Tạo trục y thứ hai ở bên phải (optional)
    ax2 = ax.twinx()  # Trục y phụ ở bên phải
    ax2.set_ylim(ax.get_ylim())  # Giới hạn trục y phụ giống như trục y chính
    # Định dạng trục y với dấu phẩy
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))

    # Lưu biểu đồ vào buffer
    buffer = BytesIO()
    plt.savefig(buffer, format="png", bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return buffer


def plot_stock_price1(stock_code, selected_date):
    # Kiểm tra xem mã cổ phiếu có tồn tại trong dữ liệu không
    if stock_code not in price_df.columns:
        print(f"Mã {stock_code} không tồn tại trong dữ liệu.")
        return None

    stock_price = price_df[stock_code].dropna()

    if stock_price.empty:
        print(f"Không có dữ liệu đủ để vẽ biểu đồ cho {stock_code}.")
        return None

    # Chuyển selected_date thành datetime
    selected_date = pd.to_datetime(selected_date)

    # Lọc dữ liệu trong vòng 1 năm, nếu có đủ dữ liệu
    start_date = selected_date - timedelta(days=365)  # 365 ngày là 1 năm
    stock_price = stock_price[(stock_price.index >= start_date) & (stock_price.index <= selected_date)]

    
    
    # Vẽ biểu đồ
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(stock_price.index, stock_price.values, linestyle='-', color='#D32F2F')  # Màu đỏ
    ax.grid(True)

    # Định dạng trục y với dấu phẩy
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))

    # Hiển thị chỉ 4 mốc thời gian: T1, T4, T8, T12
    months = pd.date_range(start=stock_price.index.min(), end=stock_price.index.max(), freq='3MS')  # 4 mốc

    # Định dạng các mốc thời gian và áp dụng 'strftime' cho từng phần tử trong 'months'
    months_str = months.strftime('%m/%Y')  # Định dạng mốc thời gian

    plt.xticks(months, months_str)  # Đặt mốc thời gian lên trục X

    # Chỉnh trục y sang bên phải
    ax.yaxis.set_label_position('right')  # Đổi vị trí nhãn trục y sang bên phải
    ax.yaxis.tick_right()  # Đưa các dấu tick của trục y sang bên phải

    # Tạo trục y thứ hai ở bên phải (optional)
    ax2 = ax.twinx()  # Trục y phụ ở bên phải
    ax2.set_ylim(ax.get_ylim())  # Giới hạn trục y phụ giống như trục y chính
    # Định dạng trục y với dấu phẩy
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{int(x):,}'))

    # Lưu biểu đồ vào buffer
    buffer = BytesIO()
    plt.tight_layout()
    plt.savefig(buffer, format="png", bbox_inches='tight')
    plt.close(fig)
    buffer.seek(0)
    return buffer

def calculate_percentage_change(price_df, selected_date, stock_code):
    date_ranges = {
        "1 ngày": 1,
        "5 ngày": 5,
        "3 tháng": 90,
        "6 tháng": 180,
        "1 năm": 365
    }

    # Đảm bảo kiểu datetime
    selected_date = pd.to_datetime(selected_date).normalize()
    price_df.index = pd.to_datetime(price_df.index).normalize()

    percentage_changes = {}

    # Kiểm tra xem stock_code có tồn tại không
    if stock_code not in price_df.columns:
        return {label: "Không có dữ liệu" for label in date_ranges}

    # Kiểm tra ngày hiện tại có giá không
    if selected_date not in price_df.index or pd.isna(price_df.loc[selected_date, stock_code]):
        return {label: "Không có dữ liệu" for label in date_ranges}

    current_price = price_df.loc[selected_date, stock_code]

    for label, num_days in date_ranges.items():
        # Tìm ngày gần nhất trước đó
        past_target_date = selected_date - pd.Timedelta(days=num_days)
        past_dates = price_df.index[(price_df.index <= past_target_date) & (~price_df[stock_code].isna())]

        if len(past_dates) == 0:
            percentage_changes[label] = "Không có dữ liệu"
            continue

        past_date = past_dates[-1]
        past_price = price_df.loc[past_date, stock_code]

        # Tính phần trăm thay đổi
        change = ((current_price - past_price) / past_price) * 100
        percentage_changes[label] = round(change, 2)

    return percentage_changes

def draw_profitability_chart(ratio_df, stock_code):
    df_plot = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))].sort_values("Năm")
    if df_plot.empty:
        return None

    buffer = BytesIO()
    plt.figure(figsize=(8, 4))

    # Chỉ định màu cho các đường
    plt.plot(df_plot["Năm"], df_plot["ROA (%)"], marker='o', label="ROA (%)", color='#4B4B4B')  # Màu xám đậm
    plt.plot(df_plot["Năm"], df_plot["ROE (%)"], marker='o', label="ROE (%)", color='#D32F2F')  # Màu đỏ
    plt.plot(df_plot["Năm"], df_plot["ROS (%)"], marker='o', label="ROS (%)", color='#D3D3D3')  # Màu xám nhạt

    plt.xticks(df_plot["Năm"].astype(int))
    plt.ylabel("Tỷ lệ (%)")
    plt.xlabel("Năm")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()
    plt.savefig(buffer, format="png", dpi=150)
    plt.close()

    buffer.seek(0)
    return buffer

def draw_valuation_chart(ratio_df, industry_avg_df, stock_code):
    df_company = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))]
    industry_name = df_company["Ngành ICB - cấp 3"].iloc[0] if not df_company.empty else None
    df_industry = industry_avg_df[
        (industry_avg_df["Ngành ICB - cấp 3"] == industry_name) &
        (industry_avg_df["Năm"].between(2020, 2024))
    ]

    if df_company.empty or df_industry.empty:
        return None

    # Chuẩn hóa lại năm
    years = sorted(df_company["Năm"].unique())
    pe_company = df_company.set_index("Năm")["P/E"]
    pb_company = df_company.set_index("Năm")["P/B"]
    pe_industry = df_industry.set_index("Năm")["P/E"]
    pb_industry = df_industry.set_index("Năm")["P/B"]

    x = range(len(years))
    width = 0.2

    # Tạo buffer để lưu ảnh
    buffer = BytesIO()
    plt.figure(figsize=(10, 5))

    # Đổi màu cho các cột
    plt.bar([i - width for i in x], [pe_company.get(y, 0) for y in years], width=width, label='P/E - Công ty', color='#4B4B4B')  # Màu xám đậm
    plt.bar(x, [pe_industry.get(y, 0) for y in years], width=width, label='P/E - TB ngành', color='#D32F2F')  # Màu đỏ

    plt.bar([i + width*1.5 for i in x], [pb_company.get(y, 0) for y in years], width=width, label='P/B - Công ty', color='#D3D3D3')  # Màu xám nhạt
    plt.bar([i + width*2.5 for i in x], [pb_industry.get(y, 0) for y in years], width=width, label='P/B - TB ngành', color='#FF6F61')  # Màu đỏ pastel

    plt.xticks(x, years)
    plt.xlabel("Năm")
    plt.legend()
    plt.tight_layout()
    plt.grid(axis='y', linestyle='--', alpha=0.5)
    plt.savefig(buffer, format="png", dpi=150)
    plt.close()
    buffer.seek(0)
    return buffer

def draw_growth_chart(ratio_df, stock_code):
    # Lọc dữ liệu theo mã và năm từ 2020 đến 2024
    df_plot = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))].sort_values("Năm")

    if df_plot.empty:
        print("Không có dữ liệu tăng trưởng.")
        return None

    buffer = BytesIO()
    plt.figure(figsize=(8, 4))

    # Vẽ đường tăng trưởng doanh thu và lợi nhuận với các màu sắc
    plt.plot(df_plot["Năm"], df_plot["Revenue Growth (%)"], marker='o', label="Tăng trưởng Doanh thu (%)", color='#4B4B4B')  # Màu xám đậm
    plt.plot(df_plot["Năm"], df_plot["Net Income Growth (%)"], marker='o', label="Tăng trưởng LNST (%)", color='#D32F2F')  # Màu đỏ

    # Định dạng biểu đồ
    plt.xticks(df_plot["Năm"].astype(int))
    plt.ylabel("Tăng trưởng (%)")
    plt.xlabel("Năm")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()

    # Lưu biểu đồ vào buffer
    plt.savefig(buffer, format="png", dpi=150)
    plt.close()
    buffer.seek(0)
    return buffer

def draw_leverage_chart(ratio_df, stock_code):
    # Lọc dữ liệu theo mã và năm từ 2020 đến 2024
    df_plot = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))].sort_values("Năm")

    if df_plot.empty:
        print("Không đủ dữ liệu để vẽ biểu đồ.")
        return None

    # Tạo buffer để lưu hình
    buffer = BytesIO()
    plt.figure(figsize=(8, 4))

    # Vẽ các đường chỉ số với các màu sắc
    plt.plot(df_plot["Năm"], df_plot["D/A (%)"], marker='o', label="D/A (%)", color='#4B4B4B')  # Màu xám đậm
    plt.plot(df_plot["Năm"], df_plot["D/E (%)"], marker='o', label="D/E (%)", color='#D32F2F')  # Màu đỏ
    plt.plot(df_plot["Năm"], df_plot["E/A (%)"], marker='o', label="E/A (%)", color='#D3D3D3')  # Màu xám nhạt

    # Tuỳ chỉnh trục và hiển thị
    plt.xticks(df_plot["Năm"].astype(int))
    plt.ylabel("Tỷ lệ (%)")
    plt.xlabel("Năm")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.legend()
    plt.tight_layout()

    # Lưu biểu đồ vào buffer để chèn vào PDF
    plt.savefig(buffer, format="png", dpi=150)
    plt.close()
    buffer.seek(0)

    return buffer


def add_page_footer(c, width):
    c.setFont("Roboto_Regular", 11)
    c.setFillColor(colors.black)
    c.drawCentredString(width / 2, 20, f"Trang {c.getPageNumber()}")

def create_pdf(stock_code, report_date):
    """Tạo file PDF chứa thông tin doanh nghiệp theo mã đã chọn."""
    stock_info = df[df["Mã"] == stock_code]
    if stock_info.empty:
        return None
  
    # Đọc KQKD từ file CSV, dùng đường dẫn tuyệt đối
    base_path = os.path.dirname(__file__)
    kqkd_file_path = os.path.join(base_path, 'KQKD.csv')

    if os.path.exists(kqkd_file_path):
        kqkd_df = pd.read_csv(kqkd_file_path)
    else:
        kqkd_df = None
        print(f"❌ Không tìm thấy file: {kqkd_file_path}")

    # Lấy thông tin
    ten_cong_ty = stock_info.iloc[0]["Tên công ty"]
    san = stock_info.iloc[0]["Sàn"]
    nganh_cap1 = stock_info.iloc[0]["Ngành ICB - cấp 1"]
    nganh_cap2 = stock_info.iloc[0]["Ngành ICB - cấp 2"]
    nganh_cap3 = stock_info.iloc[0]["Ngành ICB - cấp 3"]
    nganh_cap4 = stock_info.iloc[0]["Ngành ICB - cấp 4"]
    ngay_tao = report_date.strftime('%d-%m-%Y')

    # Lấy thông tin tóm tắt doanh nghiệp từ file info.xlsx
    info_data = info_df[info_df["Mã CK"] == stock_code]
    tom_tat = info_data["Thông tin"].values[0] if not info_data.empty else "Không có thông tin."

    # Lấy giá cổ phiếu từ price_df
    report_date = pd.to_datetime(selected_date)  # Đảm bảo kiểu datetime64
    if report_date in price_df.index and stock_code in price_df.columns:
        stock_price = price_df.loc[report_date, stock_code]
    else:
        stock_price = "Không có dữ liệu"

    # Tính giá cao nhất và thấp nhất trong 52 tuần trước ngày báo cáo
    start_date_52_weeks = report_date - timedelta(weeks=52)
    stock_price_52_weeks = price_df[stock_code].loc[start_date_52_weeks:report_date]

    highest_52_weeks = stock_price_52_weeks.max() if not stock_price_52_weeks.empty else "Không có dữ liệu"
    lowest_52_weeks = stock_price_52_weeks.min() if not stock_price_52_weeks.empty else "Không có dữ liệu"

    # Lấy thông tin SLCP lưu hành từ ratio.xlsx
    ratio_df['Năm'] = pd.to_numeric(ratio_df['Năm'], errors='coerce')
    ratio_df['Mã'] = ratio_df['Mã'].astype(str).str.strip()
    ratio_data = ratio_df[(ratio_df['Mã'] == stock_code) & (ratio_df['Năm'] == selected_date.year)]
    slcp = ratio_data['SLCP lưu hành'].values[0] if not ratio_data.empty else "Không có dữ liệu"

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    add_page_footer(c, width)
    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")
    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)

    # Tiêu đề mục
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)  
    c.drawString(40, height - 107, "GIỚI THIỆU")
    c.setFillColor(colors.black)

    # Dữ liệu bảng
    data = [
        ["Mã chứng khoán", stock_code],
        ["Tên công ty", ten_cong_ty],
        ["Sàn niêm yết", san],
        ["Ngành, phân ngành", f"{nganh_cap1} - {nganh_cap2} - {nganh_cap3} - {nganh_cap4}"]
    ]

    # Tạo bảng
    table = Table(data, colWidths=[100, width - 180])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Roboto_Black'),
        ('FONTNAME', (1, 0), (1, -1), 'Roboto_Regular'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),

        # Màu nền xen kẽ theo hàng: 
        ('BACKGROUND', (0, 1), (-1, 1), bg),  # 2
        ('BACKGROUND', (0, 3), (-1, 3), bg),  
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),  # 1
        ('BACKGROUND', (0, 2), (-1, 2), colors.white), 
    ]))

    # Vẽ bảng lên PDF
    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 205)

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 125 - len(data) * 22, width - 40, height - 125 - len(data) * 22)

    # Tiêu đề "SƠ LƯỢC VỀ DOANH NGHIỆP"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)  
    c.drawString(40, height - 145 - len(data) * 22, "SƠ LƯỢC VỀ DOANH NGHIỆP")

    c.setFillColor(colors.black)

    # Hiển thị nội dung tóm tắt doanh nghiệp
    styles = getSampleStyleSheet()
    styleN = ParagraphStyle(
        'Normal',
        parent=styles["Normal"],
        fontName="Roboto_Regular",  
        fontSize=11,
        leading=15, 
        alignment=TA_JUSTIFY
    )

    p = Paragraph(tom_tat, styleN)

    # Xác định vị trí và kích thước vùng văn bản
    w, h = p.wrap(width - 80, height - 240)
    p.drawOn(c, 40, height - 240 - h)

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 250 - h, width - 40, height - 250 - h)

    # Tiêu đề "BIỂU ĐỒ GIÁ"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)  
    c.drawString(40, height - 270 - h, "BIỂU ĐỒ GIÁ")

    c.setFillColor(colors.black)

    # Chữ "Giá cổ phiếu dài hạn"
    c.setFont("Roboto_Regular", 9)
    c.drawString(120, height - 280 - h, "Giá cổ phiếu dài hạn")

    # Vẽ biểu đồ giá cổ phiếu dài hạn
    chart_buffer = plot_stock_price(stock_code)
    if chart_buffer:
        img = Image(chart_buffer, width=240, height=140)
        img.wrapOn(c, width, height)
        img.drawOn(c, 40, height - 425 - h)

    # Chữ "Giá cổ phiếu 1 năm"
    c.setFont("Roboto_Regular", 9)
    c.drawString(395, height - 280 - h, "Giá cổ phiếu 1 năm")

    # Vẽ biểu đồ giá cổ phiếu 1 năm
    chart_buffer = plot_stock_price1(stock_code, selected_date)
    if chart_buffer:
        img = Image(chart_buffer, width=240, height=140)
        img.wrapOn(c, width, height)
        img.drawOn(c, 315, height - 425 - h)  # Vị trí biểu đồ giá cổ phiếu 1 năm

    # Line phân cách
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 435 - h, width - 310, height - 435 - h)


    # Tiêu đề "THÔNG TIN CỔ PHIẾU"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 455 - h, "THÔNG TIN CỔ PHIẾU")

    c.setFillColor(colors.black)

    # Thêm bảng thông tin nhỏ phía dưới "Thông tin chung"
    small_table_data = [
        ["Giá đóng cửa", f"{int(stock_price):,}"],
        ["Giá cao nhất 52 tuần qua", f"{int(highest_52_weeks):,}" if highest_52_weeks != "Không có dữ liệu" else highest_52_weeks],
        ["Giá thấp nhất 52 tuần qua", f"{int(lowest_52_weeks):,}" if lowest_52_weeks != "Không có dữ liệu" else lowest_52_weeks],
        ["Số lượng cổ phiếu lưu hành", f"{int(slcp):,}" if slcp != "Không có dữ liệu" else slcp],
        ["Đơn vị tiền tệ", "VND"]
    ]

    # Tạo bảng nhỏ
    small_table = Table(small_table_data, colWidths=[140, width - 490])
    small_table.setStyle(TableStyle([
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, -1), 'Roboto_Black'),
        ('FONTNAME', (1, 0), (1, -1), 'Roboto_Regular'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), bg),
        ('BACKGROUND', (0, 2), (-1, 2), colors.white),
        ('BACKGROUND', (0, 3), (-1, 3), bg),
        ('BACKGROUND', (0, 4), (-1, 4), colors.white),
    ]))

    # Vẽ bảng nhỏ lên PDF
    small_table.wrapOn(c, width, height)
    small_table.drawOn(c, 40, height - 580 - h)

    # Line nhỏ
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(320, height - 435 - h, width - 40, height - 435 - h)

    # Tiêu đề "BIẾN ĐỘNG GIÁ"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(400, height - 455 - h, "BIẾN ĐỘNG GIÁ")

    percentage_changes = calculate_percentage_change(price_df, selected_date, stock_code)

    data1 = [
        ["1 ngày", f"{percentage_changes['1 ngày']}%" if isinstance(percentage_changes["1 ngày"], (int, float)) else percentage_changes["1 ngày"]],
        ["5 ngày", f"{percentage_changes['5 ngày']}%" if isinstance(percentage_changes["5 ngày"], (int, float)) else percentage_changes["5 ngày"]],
        ["3 tháng", f"{percentage_changes['3 tháng']}%" if isinstance(percentage_changes["3 tháng"], (int, float)) else percentage_changes["3 tháng"]],
        ["6 tháng", f"{percentage_changes['6 tháng']}%" if isinstance(percentage_changes["6 tháng"], (int, float)) else percentage_changes["6 tháng"]],
        ["1 năm", f"{percentage_changes['1 năm']}%" if isinstance(percentage_changes["1 năm"], (int, float)) else percentage_changes["1 năm"]],
    ]

    table = Table(data1, colWidths=[140, width - 500])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),  
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'), 
        ('FONTNAME', (0, 0), (0, -1), 'Roboto_Black'),  
        ('FONTNAME', (1, 0), (1, -1), 'Roboto_Regular'),  
        ('FONTSIZE', (0, 0), (-1, -1), 11),  
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),  
        ('BACKGROUND', (0, 0), (-1, 0), bg), 
        ('BACKGROUND', (0, 1), (-1, 1), colors.white), 
        ('BACKGROUND', (0, 2), (-1, 2), bg), 
        ('BACKGROUND', (0, 3), (-1, 3), colors.white),  
        ('BACKGROUND', (0, 4), (-1, 4), bg), 
    ]))

    # Vẽ bảng 
    table.wrapOn(c, width, height)
    table.drawOn(c, 320, height - 580 - h)

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 595 - h, width - 310, height - 595 - h)

    # Vẽ đường kẻ xanh dương dưới đoạn văn bản
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(320, height - 595 - h, width - 40, height - 595 - h)

    # Tiêu đề "CHỈ SỐ TÀI CHÍNH QUAN TRỌNG"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)  
    c.drawString(40, height - 615 - h, "CHỈ SỐ TÀI CHÍNH QUAN TRỌNG")

    # Đọc dữ liệu ratio theo mã và năm
    ratio_row = ratio_df[(ratio_df['Mã'] == stock_code) & (ratio_df['Năm'] == selected_date.year)]
    kqkd_df = pd.read_csv(os.path.join(base_path, "KQKD.csv"))  

    # Hàm định dạng số
    def fmt(x, is_int=False):
        if isinstance(x, (int, float)):
            return f"{int(round(x)):,}" if is_int else f"{round(x, 2):,}"
        return x

    eps_value = "Không có dữ liệu"
    pe = "Không có dữ liệu"
    pb = "Không có dữ liệu"
    book_value = "Không có dữ liệu"

    # Xử lý
    if not ratio_row.empty:
        row = ratio_row.iloc[0]

        # EPS từ KQKD
        eps_row = kqkd_df[(kqkd_df['Mã'] == stock_code) & (kqkd_df['Năm'] == selected_date.year)]
        if not eps_row.empty:
            eps_value = eps_row["Lãi cơ bản trên cổ phiếu"].values[0]

        # Lấy các chỉ số khác
        pe = row.get("P/E", "Không có dữ liệu")
        pb = row.get("P/B", "Không có dữ liệu")

        # Giá trị sổ sách 
        try:
            book_value = float(row.get("Giá trị sổ sách", None))
        except:
            book_value = "Không có dữ liệu"

    # Tạo bảng hiển thị
    financial_table_data_1 = [["EPS (VND)", fmt(eps_value, is_int=True)]]
    if pe != "Không có dữ liệu":
        financial_table_data_1.append(["P/E", fmt(pe)])

    financial_table_data_2 = [
        ["Giá trị sổ sách (VND)", fmt(book_value, is_int=True)],
        ["P/B", fmt(pb)],
    ]

    table = Table(financial_table_data_1, colWidths=[140, width - 490])
    table.setStyle(TableStyle([
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, -1), 'Roboto_Black'),
        ('FONTNAME', (1, 0), (1, -1), 'Roboto_Regular'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), bg),
    ]))

    # Vẽ bảng lên PDF
    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 665 - h)

    # Tạo và vẽ bảng
    table2 = Table(financial_table_data_2, colWidths=[140, width - 490])
    table2.setStyle(TableStyle([
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (0, -1), 'Roboto_Black'),
        ('FONTNAME', (1, 0), (1, -1), 'Roboto_Regular'),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), bg),
    ]))

    # Vẽ bảng lên PDF
    table2.wrapOn(c, width, height)
    table2.drawOn(c, 320, height - 665 - h)

    # Nextpage
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Kết quả kinh doanh"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "KẾT QUẢ KINH DOANH")

    def plot_kqkd_yearly_to_buffer(stock_code):
        from io import BytesIO
        import matplotlib.pyplot as plt

        # ✅ Khai báo màu sắc chuẩn theo tone
        red_text = "#D32F2F"      # đỏ đậm
        red_light = "#FF8A80"     # đỏ nhạt
        bg = "#D3D3D3"            # xám nhạt
        grey = "#4B4B4B"          # xám đậm

        df = kqkd_df.copy()

        # Chuẩn hóa cột
        df.rename(columns={
            "Mã": "MÃ",
            "Năm": "NĂM",
            "Doanh thu thuần": "DOANH THU THUẦN",
            "Lợi nhuận gộp về bán hàng và cung cấp dịch vụ": "LỢI NHUẬN GỘP",
            "Lợi nhuận sau thuế thu nhập doanh nghiệp": "LNST"
        }, inplace=True)

        stock_code = stock_code.upper()
        filtered_df = df[df["MÃ"] == stock_code].copy()
        filtered_df = filtered_df.dropna(subset=["NĂM", "DOANH THU THUẦN", "LỢI NHUẬN GỘP", "LNST"])
        filtered_df["NĂM"] = filtered_df["NĂM"].astype(int)

        if filtered_df.empty:
            print(f"Không có dữ liệu đủ cho mã: {stock_code}")
            return None

        # Gộp theo năm
        yearly_df = filtered_df.groupby("NĂM")[["DOANH THU THUẦN", "LỢI NHUẬN GỘP", "LNST"]].sum().reset_index()
        yearly_df["BIÊN LỢI NHUẬN GỘP"] = yearly_df["LỢI NHUẬN GỘP"] / yearly_df["DOANH THU THUẦN"] * 100
        yearly_df["BIÊN LNST"] = yearly_df["LNST"] / yearly_df["DOANH THU THUẦN"] * 100

        # Đơn vị: tỷ đồng
        yearly_df[["DOANH THU THUẦN", "LỢI NHUẬN GỘP", "LNST"]] /= 1e9

        # Vẽ biểu đồ
        fig, axs = plt.subplots(2, 1, figsize=(12, 8), gridspec_kw={'height_ratios': [1, 1.2]})
        fig.suptitle(f"Kết quả kinh doanh - {stock_code}", fontsize=16, color=red_text)

        axs[0].bar(yearly_df["NĂM"].astype(str), yearly_df["DOANH THU THUẦN"], color=bg)
        axs[0].set_title("Doanh thu thuần theo năm", fontsize=13)
        axs[0].set_ylabel("Tỷ đồng")
        axs[0].grid(axis="y", linestyle="--", alpha=0.5)

        max_val = yearly_df["DOANH THU THUẦN"].max()
        min_val = yearly_df["DOANH THU THUẦN"].min()
        mean_val = yearly_df["DOANH THU THUẦN"].mean()
        text = f"Cao nhất: {max_val:,.0f} tỷ\nTrung bình: {mean_val:,.0f} tỷ\nThấp nhất: {min_val:,.0f} tỷ"
        axs[0].annotate(text, xy=(1.01, 0.8), xycoords='axes fraction', fontsize=10, ha="left", va="top")

        bar1 = axs[1].bar(yearly_df["NĂM"].astype(str), yearly_df["LỢI NHUẬN GỘP"], color=grey, label="Lợi nhuận gộp (tỷ đồng)")
        bar2 = axs[1].bar(yearly_df["NĂM"].astype(str), yearly_df["LNST"], color=bg, label="LNST (tỷ đồng)")
        axs[1].set_ylabel("Tỷ đồng")
        axs[1].grid(axis="y", linestyle="--", alpha=0.5)

        ax2 = axs[1].twinx()
        ax2.plot(yearly_df["NĂM"].astype(str), yearly_df["BIÊN LỢI NHUẬN GỘP"], color=red_text, label="Biên LN gộp (%)", linewidth=2)
        ax2.plot(yearly_df["NĂM"].astype(str), yearly_df["BIÊN LNST"], color=red_light, label="Biên LNST (%)", linewidth=2)
        ax2.set_ylabel("Biên lợi nhuận (%)")
        ax2.set_ylim(0, 50)

        lines_labels = ax2.lines
        labels = [bar.get_label() for bar in [bar1, bar2]] + [line.get_label() for line in lines_labels]
        axs[1].legend([bar1, bar2] + list(lines_labels), labels, loc="upper left")

        plt.tight_layout(rect=[0, 0, 1, 0.95])

        buffer = BytesIO()
        plt.savefig(buffer, format="png", bbox_inches="tight")
        plt.close(fig)
        buffer.seek(0)
        buffer.name = f"{stock_code}_kqkd_nam.png"
        return buffer
    chart_buffer = plot_kqkd_yearly_to_buffer(stock_code)
    if chart_buffer:
        img = Image(chart_buffer, width=520, height=310)
        img.wrapOn(c, width, height)
        img.drawOn(c, 60, height - 425)
    # Vẽ lại biểu đồ 4 chỉ tiêu KQKD theo style PDF hiện tại (giao diện màu đỏ-xám, đơn giản, rõ ràng)
    def draw_kqkd_summary_chart(kqkd_df, stock_code):
        df = kqkd_df.copy()
        df = df[df["Mã"] == stock_code.upper()]
        df = df.dropna(subset=[
            "Doanh thu thuần",
            "Tổng lợi nhuận kế toán trước thuế",
            "Lợi nhuận sau thuế thu nhập doanh nghiệp"
        ])
        df["Năm"] = df["Năm"].astype(int)
        df = df[df["Năm"].between(2020, 2024)]
        df_grouped = df.groupby("Năm").agg({
            "Doanh thu thuần": "sum",
            "Tổng lợi nhuận kế toán trước thuế": "sum",
            "Lợi nhuận sau thuế thu nhập doanh nghiệp": "sum"
        }).reset_index()
        df_grouped["Biên LN ròng (%)"] = df_grouped["Lợi nhuận sau thuế thu nhập doanh nghiệp"] / df_grouped["Doanh thu thuần"] * 100

        # Đơn vị: tỷ
        for col in ["Doanh thu thuần", "Tổng lợi nhuận kế toán trước thuế", "Lợi nhuận sau thuế thu nhập doanh nghiệp"]:
            df_grouped[col] /= 1e9

        # Màu
        grey = "#4B4B4B"

        # Vẽ
        fig, axs = plt.subplots(2, 2, figsize=(10, 6.5))
        fig.suptitle(f"Tổng quan KQKD - {stock_code.upper()} (2020–2024)", fontsize=14, color=grey)

        axs[0, 0].bar(df_grouped["Năm"], df_grouped["Doanh thu thuần"], color=grey)
        axs[0, 0].set_title("Doanh thu", fontsize=11)
        axs[0, 0].set_ylabel("Tỷ đồng")
        
        axs[0, 1].bar(df_grouped["Năm"], df_grouped["Tổng lợi nhuận kế toán trước thuế"], color=grey)
        axs[0, 1].set_title("LN trước thuế", fontsize=11)
        axs[0, 1].set_ylabel("Tỷ đồng")

        axs[1, 0].bar(df_grouped["Năm"], df_grouped["Lợi nhuận sau thuế thu nhập doanh nghiệp"], color=grey)
        axs[1, 0].set_title("LN sau thuế", fontsize=11)
        axs[1, 0].set_ylabel("Tỷ đồng")

        axs[1, 1].bar(df_grouped["Năm"], df_grouped["Biên LN ròng (%)"], color=grey)
        axs[1, 1].set_title("Biên lợi nhuận ròng (%)", fontsize=11)
        axs[1, 1].set_ylabel("%")

        for ax in axs.flat:
            ax.set_xticks(df_grouped["Năm"])
            ax.grid(axis="y", linestyle="--", alpha=0.4)

        plt.tight_layout(rect=[0, 0, 1, 0.95])

        # Export to buffer
        buf = BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        buf.name = f"{stock_code}_kqkd_summary.png"
        return buf
    chart_buffer = draw_kqkd_summary_chart(kqkd_df, stock_code)
    if chart_buffer:
        img = Image(chart_buffer, width=520, height=300)
        img.wrapOn(c, width, height)
        img.drawOn(c, 40, height - 740) 
    # Nextpage
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Báo cáo tài chính"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "BÁO CÁO TÀI CHÍNH")

    c.setFillColor(colors.black)

    #Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 112, 160, height - 112)

    # Tiêu đề "Bảng cân đối kế toán"
    c.setFont("Roboto_Black", 14)
    c.drawString(40, height - 130, "Bảng cân đối kế toán")


    # Lọc theo mã
    bcdkt_stock = bcdkt_df[bcdkt_df['Mã'] == stock_code]

    # Lấy các năm theo thứ tự tăng dần
    years = sorted(bcdkt_stock['Năm'].dropna().astype(int).unique())

    # Tạo tiêu đề bảng
    headers = ["Chỉ tiêu"] + [str(y) for y in years]

    # Các chỉ tiêu cần hiển thị
    fields = [
        "TÀI SẢN NGẮN HẠN",
        "TÀI SẢN DÀI HẠN",
        "TỔNG CỘNG TÀI SẢN",
        "NỢ PHẢI TRẢ",
        "VỐN CHỦ SỞ HỮU",
        "TỔNG CỘNG NGUỒN VỐN",
    ]

    # Tạo dữ liệu bảng
    data = [headers]
    for field in fields:
        row = [field.replace("_", " ").title()]
        for y in years:
            val = bcdkt_stock.loc[(bcdkt_stock['Năm'] == y), field]
            if not val.empty and pd.notna(val.values[0]):
                value_million = int(val.values[0]) // 1_000_000
                row.append(f"{value_million:,}")
            else:
                row.append("Không có")
        data.append(row)

    # Tạo bảng PDF
    usable_width = width - 80
    colWidths = [160] + [(usable_width - 160) / len(years)] * len(years)
    table = Table(data, colWidths=colWidths)

    # Tạo danh sách dòng có nền xen kẽ (bỏ dòng đầu vì là header)
    background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]  

    for i in range(1, len(data)):
        bg_color = bg if i % 2 == 1 else colors.white
        background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

    # Áp dụng toàn bộ style
    table.setStyle(TableStyle([
                                  ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                                  ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                  ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),  # Header
                                  ('FONTNAME', (0, 1), (0, -1), 'Roboto_Regular'),  # Tên chỉ tiêu
                                  ('FONTNAME', (1, 1), (-1, -1), 'Roboto_Regular'),  # Dữ liệu
                                  ('FONTSIZE', (0, 0), (-1, -1), 10),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                  ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ] + background_styles))

    # Vẽ bảng chính giữa
    table.wrapOn(c, width, height)
    x_pos = (width - sum(colWidths)) / 2
    table.drawOn(c, x_pos, height - 305)

    c.setFont("Roboto_Regular", 10)
    c.drawString(165, height - 130, "(Đơn vị: triệu VND)")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 310, width - 40, height - 310)

    # Tiêu đề "Bảng kết quả kinh doanh"
    c.setFont("Roboto_Black", 14)
    c.drawString(40, height - 330, "Bảng kết quả kinh doanh")

    c.setFont("Roboto_Regular", 10)
    c.drawString(185, height - 330, "(Đơn vị: triệu VND)")

    # Lọc theo mã
    kqkd_stock = kqkd_df[kqkd_df['Mã'] == stock_code]

    # Lấy các năm theo thứ tự
    years = sorted(kqkd_stock['Năm'].dropna().astype(int).unique())

    # Tiêu đề
    headers = ["Chỉ tiêu"] + [str(y) for y in years]

    # Các chỉ tiêu cần hiển thị
    fields = [
        "Doanh thu thuần",
        "Lợi nhuận thuần từ hoạt động kinh doanh",
        "Tổng lợi nhuận kế toán trước thuế",
        "Lợi nhuận sau thuế thu nhập doanh nghiệp",
        "Lãi trước thuế"
    ]

    # Tạo dữ liệu bảng
    data = [headers]
    for field in fields:
        # Nếu cần đổi tên hiển thị
        display_name = (
            "Lợi nhuận sau thuế" if field == "Lợi nhuận sau thuế thu nhập doanh nghiệp" else field
        )
        row = [display_name]
        for y in years:
            val = kqkd_stock.loc[(kqkd_stock['Năm'] == y), field]
            if not val.empty and pd.notna(val.values[0]):
                value = int(val.values[0]) // 1_000_000  # Chuyển về triệu nếu cần
                row.append(f"{value:,}")
            else:
                row.append("Không có")
        data.append(row)

    usable_width = width - 80
    colWidths = [160] + [(usable_width - 160) / len(years)] * len(years)

    table = Table(data, colWidths=colWidths)

    # Màu nền xen kẽ
    background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]
    for i in range(1, len(data)):
        bg_color = bg if i % 2 == 1 else colors.white
        background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

    # Style bảng
    table.setStyle(TableStyle([
                                  ('ALIGN', (1, 1), (len(years), -1), 'RIGHT'),
                                  ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                  ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                  ('FONTNAME', (0, 1), (0, -1), 'Roboto_Regular'),
                                  ('FONTNAME', (1, 1), (-1, -1), 'Roboto_Regular'),
                                  ('FONTSIZE', (0, 0), (-1, -1), 10),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                  ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ] + background_styles))

    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 480)

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 490, width - 40, height - 490)

    # Tiêu đề "Bảng lưu chuyển tiền tệ"
    c.setFont("Roboto_Black", 14)
    c.drawString(40, height - 510, "Bảng lưu chuyển tiền tệ")

    c.setFont("Roboto_Regular", 10)
    c.drawString(180, height - 510, "(Đơn vị: triệu VND)")

    # Lọc theo mã
    lctt_stock = lctt_df[lctt_df['Mã'] == stock_code]
    years = sorted(lctt_stock['Năm'].dropna().astype(int).unique())

    # Header bảng
    headers = ["Chỉ tiêu"] + [str(y) for y in years]

    # Các chỉ tiêu cần lấy và tên hiển thị
    field_map = {
        "Cổ tức đã trả (TT)": "Cổ tức đã trả",
        "Lưu chuyển tiền thuần trong kỳ (TT)": "Lưu chuyển tiền thuần",
        "Tiền và tương đương tiền đầu kỳ (TT)": "Tiền và tương đương tiền đầu kỳ",
        "Tiền và tương đương tiền cuối kỳ (TT)": "Tiền và tương đương tiền cuối kỳ"
    }

    # Tạo bảng dữ liệu
    data = [headers]
    for field, label in field_map.items():
        row = [label]
        for y in years:
            val = lctt_stock.loc[(lctt_stock['Năm'] == y), field]
            if not val.empty and pd.notna(val.values[0]):
                value = int(val.values[0]) // 1_000_000
                row.append(f"{value:,}")
            else:
                row.append("Không có")
        data.append(row)

    usable_width = width - 80
    colWidths = [160] + [(usable_width - 160) / len(years)] * len(years)

    table = Table(data, colWidths=colWidths)

    # Màu nền xen kẽ
    background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]
    for i in range(1, len(data)):
        bg_color = bg if i % 2 == 1 else colors.white
        background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

    # Style bảng
    table.setStyle(TableStyle([
                                  ('ALIGN', (1, 1), (len(years), -1), 'CENTER'),
                                  ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                  ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                  ('FONTNAME', (0, 1), (0, -1), 'Roboto_Regular'),
                                  ('FONTNAME', (1, 1), (-1, -1), 'Roboto_Regular'),
                                  ('FONTSIZE', (0, 0), (-1, -1), 10),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                  ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ] + background_styles))

    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 640)

    # Ngắt sang trang mới
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)

    # Tiêu đề "Các chỉ số tài chính"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "CÁC CHỈ SỐ TÀI CHÍNH")

    # Tiêu đề "Chỉ số định giá (Valuation Ratios)"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 127, "1. Chỉ số định giá (Valuation Ratios)")

    c.setFillColor(colors.black)

    # Đọc dữ liệu ngành
    industry_avg_df = pd.read_excel(os.path.join(base_path, "ratio.xlsx"))
    industry_row = industry_avg_df[
        (industry_avg_df['Ngành ICB - cấp 3'] == nganh_cap3) & (industry_avg_df['Năm'] == selected_date.year)
        ]

    # Lấy P/E và P/B của công ty
    pe_value = ratio_row.iloc[0].get('P/E', None)
    pb_value = ratio_row.iloc[0].get('P/B', None)

    # Lấy P/E và P/B của ngành
    pe_ind = industry_row["P/E"].values[0] if not industry_row.empty else "NA"
    pb_ind = industry_row["P/B"].values[0] if not industry_row.empty else "NA"

    # So sánh bằng biểu tượng
    def compare_icon(val1, val2):
        if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
            if val1 > val2:
                return Image(os.path.join(base_path, "arrow_up_green.png"), width=10, height=10)
            elif val1 < val2:
                return Image(os.path.join(base_path, "arrow_down_red.png"), width=10, height=10)
        return "-"

    # Tạo bảng dữ liệu
    valuation_data = [
        ["Chỉ số", ten_cong_ty, "Trung bình ngành", "So sánh"],
        ["P/E", f"{round(pe_value, 2):,}" if isinstance(pe_value, (int, float)) else "NA",
         f"{round(pe_ind, 2):,}" if isinstance(pe_ind, (int, float)) else "NA",
         compare_icon(pe_value, pe_ind)],
        ["P/B", f"{round(pb_value, 2):,}" if isinstance(pb_value, (int, float)) else "NA",
         f"{round(pb_ind, 2):,}" if isinstance(pb_ind, (int, float)) else "NA",
         compare_icon(pb_value, pb_ind)],
    ]

    valuation_table = Table(valuation_data, colWidths=[130, 130, 130, 125])

    # Màu nền xen kẽ
    valuation_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]
    for i in range(1, len(valuation_data)):
        bg_color = bg if i % 2 == 1 else colors.white
        valuation_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

    # Áp dụng TableStyle
    valuation_table.setStyle(TableStyle([
                                            ('ALIGN', (1, 1), (2, -1), 'LEFT'),
                                            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                            ('ALIGN', (3, 1), (3, -1), 'LEFT'),
                                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                                            ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                            ('FONTNAME', (0, 1), (-1, -1), 'Roboto_Regular'),
                                            ('FONTSIZE', (0, 0), (-1, -1), 10),
                                            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                            ('TOPPADDING', (0, 0), (-1, -1), 6),
                                        ] + valuation_styles))

    # Vẽ bảng
    valuation_table.wrapOn(c, width, height)
    valuation_table.drawOn(c, 40, height - 210)

    # Tiêu đề "Biểu đồ so sánh"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 230, "Biểu đồ so sánh")

    # Vẽ biểu đồ
    chart_buffer = draw_valuation_chart(ratio_df, pd.read_excel(os.path.join(base_path,"industry_avg.xlsx")), stock_code) 
    if chart_buffer:
        chart_img = Image(chart_buffer, width=500, height=300)
        chart_img.wrapOn(c, width, height)
        chart_img.drawOn(c, 40, height - 545)

    c.setFillColor(colors.black)

    # Tiêu đề "Nhận xét"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 565, "Nhận xét")

    # ----- NHẬN XÉT P/E VÀ P/B -----
    pe_diff = pe_value - pe_ind if isinstance(pe_value, (int, float)) and isinstance(pe_ind, (int, float)) else None
    pb_diff = pb_value - pb_ind if isinstance(pb_value, (int, float)) and isinstance(pb_ind, (int, float)) else None

    valuation_note = ""

    if isinstance(pe_value, (int, float)) and isinstance(pe_ind, (int, float)):
        pe_diff = pe_value - pe_ind
        if pe_diff > 0:
            valuation_note += f"- Chỉ số P/E của công ty đang cao hơn trung bình ngành khoảng {pe_diff:.2f} lần. Điều này có thể cho thấy nhà đầu tư đang kỳ vọng vào tiềm năng tăng trưởng trong tương lai, hoặc cổ phiếu đang bị định giá cao so với lợi nhuận hiện tại.\n"
        elif pe_diff < 0:
            valuation_note += f"- Chỉ số P/E thấp hơn mức trung bình ngành khoảng {abs(pe_diff):.2f} lần. Đây có thể là dấu hiệu của mức giá hợp lý hoặc do thị trường đánh giá thấp khả năng sinh lời trong tương lai của công ty.\n"
        else:
            valuation_note += "- Chỉ số P/E của công ty gần như tương đương với trung bình ngành, phản ánh mức định giá ổn định theo mặt bằng chung.\n"
    else:
        valuation_note += "- Không có đủ dữ liệu để đánh giá chỉ số P/E của công ty so với ngành.\n"

    if isinstance(pb_value, (int, float)) and isinstance(pb_ind, (int, float)):
        pb_diff = pb_value - pb_ind
        if pb_diff > 0:
            valuation_note += f"- Chỉ số P/B cao hơn trung bình ngành khoảng {pb_diff:.2f} lần, cho thấy thị trường có thể đang đánh giá cao tài sản vô hình hoặc khả năng sinh lợi trong tương lai của doanh nghiệp.\n"
        elif pb_diff < 0:
            valuation_note += f"- Chỉ số P/B thấp hơn trung bình ngành khoảng {abs(pb_diff):.2f} lần, điều này có thể phản ánh sự dè dặt của thị trường hoặc dấu hiệu tiềm ẩn về hiệu quả sử dụng tài sản.\n"
        else:
            valuation_note += "- Chỉ số P/B của công ty xấp xỉ mức trung bình ngành.\n"
    else:
        valuation_note += "- Không có đủ dữ liệu để so sánh chỉ số P/B.\n"

    # Vẽ đoạn nhận xét ra PDF
    styleN = ParagraphStyle(
        'Normal',
        fontName="Roboto_Regular",
        fontSize=11,
        leading=15,
        alignment=TA_JUSTIFY,
    )

    p = Paragraph(valuation_note.replace("\n", "<br/>"), styleN)
    w, h = p.wrap(width - 80, height)
    p.drawOn(c, 40, height - 575 - h)

    # Ngắt sang trang mới
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Các chỉ số tài chính"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "CÁC CHỈ SỐ TÀI CHÍNH")

    # Tiêu đề "Chỉ số sinh lời (Profitability Ratios)"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 127, "2. Chỉ số sinh lời (Profitability Ratios)")

    c.setFillColor(colors.black)

    # Lấy dữ liệu cho ROA, ROE, ROS năm 2024
    company_row = ratio_df[(ratio_df['Mã'] == stock_code) & (ratio_df['Năm'] == 2024)]
    if not company_row.empty:
        industry_name = company_row.iloc[0]["Ngành ICB - cấp 3"]
        industry_avg_df = pd.read_excel(os.path.join(base_path, "industry_avg.xlsx"))
        industry_row = industry_avg_df[
            (industry_avg_df['Ngành ICB - cấp 3'] == industry_name) & (industry_avg_df['Năm'] == 2024)]

        if not industry_row.empty:
            def extract(val):
                return round(val, 2) if pd.notna(val) else "NA"

            def compare(val1, val2):
                if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                    return "↑" if val1 > val2 else ("↓" if val1 < val2 else "=")
                return "-"

            roe_c = extract(company_row["ROE (%)"].values[0])
            roa_c = extract(company_row["ROA (%)"].values[0])
            ros_c = extract(company_row["ROS (%)"].values[0])

            roe_i = extract(industry_row["ROE (%)"].values[0])
            roa_i = extract(industry_row["ROA (%)"].values[0])
            ros_i = extract(industry_row["ROS (%)"].values[0])

            def compare_icon_img(val1, val2):
                if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                    if val1 > val2:
                        return Image(os.path.join(base_path, "arrow_up_green.png"), width=10, height=10)
                    elif val1 < val2:
                        return Image(os.path.join(base_path, "arrow_down_red.png"), width=10, height=10)
                return "-"

            data = [
                ["Chỉ số", ten_cong_ty, "Trung bình ngành", "So sánh"],
                ["ROE (%)", roe_c, roe_i, compare_icon_img(roe_c, roe_i)],
                ["ROA (%)", roa_c, roa_i, compare_icon_img(roa_c, roa_i)],
                ["ROS (%)", ros_c, ros_i, compare_icon_img(ros_c, ros_i)],
            ]

            table1 = Table(data, colWidths=[130, 130, 130, 125])

            # Style nền xen kẽ cho bảng ROA/ROE/ROS
            background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]

            for i in range(1, len(data)):
                bg_color = bg if i % 2 == 1 else colors.white
                background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

            # Áp dụng style
            table1.setStyle(TableStyle([
                                           # Căn lề
                                           ('ALIGN', (1, 1), (2, -1), 'LEFT'),  # Cột Công ty & TB Ngành
                                           ('ALIGN', (3, 1), (3, -1), 'LEFT'),  # Cột So sánh
                                           ('ALIGN', (0, 1), (0, -1), 'LEFT'),  # Cột Chỉ số
                                           ('ALIGN', (0, 0), (-1, 0), 'LEFT'),  # Header

                                           ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                           ('FONTNAME', (0, 1), (0, -1), 'Roboto_Regular'),
                                           ('FONTNAME', (1, 1), (2, -1), 'Roboto_Regular'),
                                           ('FONTSIZE', (0, 0), (-1, -1), 10),
                                           ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                           ('TOPPADDING', (0, 0), (-1, -1), 6),
                                       ] + background_styles))

            # Vẽ bảng
            table1.wrapOn(c, width, height)
            table1.drawOn(c, 40, height - 230)

    # Tiêu đề "Biểu đồ so sánh"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 260, "Biểu đồ so sánh")

    # Gọi hàm tạo biểu đồ
    chart_buffer = draw_profitability_chart(ratio_df, stock_code)

    # Chèn vào PDF
    if chart_buffer:
        c.setFont("Roboto_Black", 14)
        c.setFillColor(bg)

        chart_img = Image(chart_buffer, width=520, height=310)
        chart_img.wrapOn(c, width, height)
        chart_img.drawOn(c, 40, height - 575)

    c.setFillColor(colors.black)

    # Tiêu đề "Nhận xét"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 565, "Nhận xét")

    # Nhận xét tự động chi tiết
    if not company_row.empty and not industry_row.empty:
        roa_diff = roa_c - roa_i
        roe_diff = roe_c - roe_i
        ros_diff = ros_c - ros_i

        comment = "Các chỉ số sinh lời (ROA, ROE, ROS) phản ánh hiệu quả hoạt động của doanh nghiệp trong việc sử dụng tài sản, vốn chủ sở hữu và khả năng sinh lợi từ doanh thu. Dưới đây là phần phân tích chi tiết:<br/>"

        # ROA
        if isinstance(roa_diff, (int, float)):
            if roa_diff > 0:
                comment += f"- ROA (Tỷ suất lợi nhuận trên tổng tài sản) của công ty đạt {roa_c}%, cao hơn mức trung bình ngành là {roa_i}%. Điều này cho thấy công ty đang sử dụng tài sản một cách hiệu quả để tạo ra lợi nhuận. Đây là dấu hiệu tích cực, thể hiện năng lực vận hành ổn định và có thể là kết quả của việc tối ưu chi phí hoạt động hoặc cấu trúc tài sản hợp lý.<br/>"
            elif roa_diff < 0:
                comment += f"- ROA của công ty chỉ đạt {roa_c}%, thấp hơn trung bình ngành là {roa_i}%. Điều này cho thấy hiệu quả sử dụng tài sản chưa tối ưu. Công ty có thể cần đánh giá lại cơ cấu tài sản, hoặc xem xét lại hoạt động vận hành để nâng cao hiệu quả sử dụng nguồn lực hiện có.<br/>"
            else:
                comment += "- ROA của công ty tương đương với trung bình ngành, cho thấy hiệu quả sử dụng tài sản ở mức trung bình so với đối thủ cạnh tranh.<br/>"

        # ROE
        if isinstance(roe_diff, (int, float)):
            if roe_diff > 0:
                comment += f"- ROE (Tỷ suất lợi nhuận trên vốn chủ sở hữu) của công ty đạt {roe_c}%, vượt trung bình ngành ({roe_i}%). Điều này chứng tỏ công ty có khả năng tạo ra giá trị cao cho cổ đông từ vốn đầu tư. Đây là một điểm mạnh cần duy trì và phát huy, đặc biệt trong việc thu hút nhà đầu tư.<br/>"
            elif roe_diff < 0:
                comment += f"- ROE của công ty là {roe_c}%, thấp hơn mức trung bình ngành là {roe_i}%. Điều này phản ánh khả năng tạo lợi nhuận từ vốn chủ sở hữu chưa hiệu quả, có thể do lợi nhuận ròng thấp hoặc vốn đầu tư chưa được khai thác đúng cách. Doanh nghiệp nên xem xét lại chiến lược sử dụng vốn hoặc cơ cấu tài chính.<br/>"
            else:
                comment += "- ROE của công ty tương đương trung bình ngành, thể hiện hiệu suất sinh lời trên vốn đầu tư ở mức phổ biến trong ngành.<br/>"

        # ROS
        if isinstance(ros_diff, (int, float)):
            if ros_diff > 0:
                comment += f"- ROS (Tỷ suất lợi nhuận trên doanh thu) của công ty là {ros_c}%, cao hơn trung bình ngành ({ros_i}%). Điều này thể hiện khả năng kiểm soát chi phí tốt và tạo ra lợi nhuận cao từ doanh thu thuần. Đây là một lợi thế cạnh tranh trong ngành có biên lợi nhuận thấp.<br/>"
            elif ros_diff < 0:
                comment += f"- ROS chỉ đạt {ros_c}%, thấp hơn trung bình ngành ({ros_i}%). Điều này có thể cho thấy công ty đang đối mặt với áp lực chi phí cao hoặc không tận dụng được lợi thế về giá bán. Cần đánh giá lại chiến lược chi phí, định giá và cấu trúc sản phẩm.<br/>"
            else:
                comment += "- ROS của công ty ngang bằng với trung bình ngành, cho thấy biên lợi nhuận ròng ở mức trung bình so với các đối thủ.<br/>"

        comment += "Tóm lại, việc so sánh các chỉ số sinh lời với trung bình ngành giúp đánh giá vị thế cạnh tranh của doanh nghiệp. Nếu các chỉ số cao hơn, công ty có lợi thế về hiệu quả và năng lực sinh lời. Ngược lại, nếu thấp hơn, cần xem xét chiến lược quản trị tài sản, chi phí và vốn để cải thiện hiệu quả hoạt động."

        # Hiển thị đoạn nhận xét lên PDF
        style_comment = ParagraphStyle(
            name="Justify",
            fontName="Roboto_Regular",
            fontSize=11,
            leading=15,
            alignment=TA_JUSTIFY,
        )

        comment_paragraph = Paragraph(comment, style_comment)
        w, h_comment = comment_paragraph.wrap(width - 80, height)
        comment_paragraph.drawOn(c, 40, height - 575 - h_comment)

    # Ngắt sang trang mới
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Các chỉ số tài chính"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "CÁC CHỈ SỐ TÀI CHÍNH")

    # Tiêu đề "3. Chỉ số tăng trưởng (Growth Ratios)"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 127, "3. Chỉ số đòn bẩy tài chính (Leverage Ratios)")

    c.setFillColor(colors.black)

    # Lấy dữ liệu
    da_c = extract(company_row["D/A (%)"].values[0])
    de_c = extract(company_row["D/E (%)"].values[0])
    ea_c = extract(company_row["E/A (%)"].values[0])

    da_i = extract(industry_row["D/A (%)"].values[0])
    de_i = extract(industry_row["D/E (%)"].values[0])
    ea_i = extract(industry_row["E/A (%)"].values[0])

    def compare_icon(val1, val2):
        if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
            if val1 > val2:
                return Image(os.path.join(base_path, "arrow_up_green.png"), width=10, height=10)
            elif val1 < val2:
                return Image(os.path.join(base_path, "arrow_down_red.png"), width=10, height=10)
        return "-"

    data_leverage = [
        ["Chỉ số", ten_cong_ty, "Trung bình ngành", "So sánh"],
        ["D/A (%)", da_c, da_i, compare_icon(da_c, da_i)],
        ["D/E (%)", de_c, de_i, compare_icon(de_c, de_i)],
        ["E/A (%)", ea_c, ea_i, compare_icon(ea_c, ea_i)],
    ]

    # Nền xen kẽ
    background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]
    for i in range(1, len(data_leverage)):
        bg_color = bg if i % 2 == 1 else colors.white
        background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

    # Tạo bảng
    table = Table(data_leverage, colWidths=[130, 130, 130, 125])
    table.setStyle(TableStyle([
                                  ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                                  ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                  ('ALIGN', (1, 1), (2, -1), 'LEFT'),
                                  ('ALIGN', (3, 1), (3, -1), 'LEFT'),
                                  ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                  ('FONTNAME', (0, 1), (-1, -1), 'Roboto_Regular'),
                                  ('FONTSIZE', (0, 0), (-1, -1), 10),
                                  ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                  ('TOPPADDING', (0, 0), (-1, -1), 6),
                              ] + background_styles))

    table.wrapOn(c, width, height)
    table.drawOn(c, 40, height - 210)

    # Tiêu đề "Biểu đồ so sánh"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 230, "Biểu đồ so sánh")

    chart_buffer = draw_leverage_chart(ratio_df, stock_code)
    if chart_buffer:
        leverage_chart = Image(chart_buffer, width=520, height=300)
        leverage_chart.wrapOn(c, width, height)
        leverage_chart.drawOn(c, 40, height - 540)

    # Tiêu đề "Nhận xét"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 550, "Nhận xét")

    def generate_leverage_comment(ratio_df, stock_code):
        df_plot = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))].sort_values("Năm")
        if df_plot.empty:
            return "Không đủ dữ liệu để nhận xét về các chỉ số đòn bẩy tài chính."

        da_series = df_plot["D/A (%)"].dropna()
        de_series = df_plot["D/E (%)"].dropna()
        ea_series = df_plot["E/A (%)"].dropna()

        avg_da = da_series.mean() if not da_series.empty else None
        avg_de = de_series.mean() if not de_series.empty else None
        avg_ea = ea_series.mean() if not ea_series.empty else None

        comment = "Phân tích các chỉ số đòn bẩy tài chính giai đoạn 2020–2024:\n"

        # Nhận xét D/A (%)
        if avg_da is not None:
            comment += f"- Tỷ lệ D/A trung bình khoảng {avg_da:.2f}%, phản ánh tỷ trọng nợ trong tổng tài sản của công ty "
            if avg_da > 50:
                comment += "ở mức cao, cho thấy công ty đang dựa nhiều vào nợ để tài trợ hoạt động, tiềm ẩn rủi ro tài chính nếu thị trường biến động mạnh.\n"
            elif avg_da >= 30:
                comment += "ở mức tương đối, cho thấy công ty đang cân bằng giữa vốn chủ sở hữu và nợ vay.\n"
            else:
                comment += "ở mức thấp, cho thấy công ty chủ yếu tài trợ bằng vốn chủ sở hữu, ít phụ thuộc vào nợ.\n"
        else:
            comment += "- Không đủ dữ liệu để đánh giá tỷ lệ D/A.\n"

        # Nhận xét D/E (%)
        if avg_de is not None:
            comment += f"- Tỷ lệ D/E trung bình khoảng {avg_de:.2f}%, phản ánh mức độ đòn bẩy tài chính của công ty. "
            if avg_de > 150:
                comment += "Tỷ lệ này khá cao, cho thấy công ty có thể gặp áp lực trả nợ lớn.\n"
            elif avg_de >= 80:
                comment += "Tỷ lệ này ở mức chấp nhận được, thể hiện công ty có sử dụng đòn bẩy nhưng chưa vượt ngưỡng rủi ro.\n"
            else:
                comment += "Tỷ lệ khá thấp, cho thấy công ty thận trọng trong vay nợ và ưu tiên vốn chủ sở hữu.\n"
        else:
            comment += "- Không đủ dữ liệu để đánh giá tỷ lệ D/E.\n"

        # Nhận xét E/A (%)
        if avg_ea is not None:
            comment += f"- Tỷ lệ E/A trung bình là {avg_ea:.2f}%, thể hiện tỷ trọng vốn chủ sở hữu trong tổng tài sản. "
            if avg_ea > 60:
                comment += "Tỷ lệ này cao cho thấy công ty có nền tảng tài chính vững chắc, ít phụ thuộc vào nợ.\n"
            elif avg_ea >= 40:
                comment += "Tỷ lệ ổn định, phản ánh sự cân đối trong cấu trúc vốn.\n"
            else:
                comment += "Tỷ lệ thấp, có thể khiến công ty gặp khó khăn khi huy động vốn trong điều kiện thị trường xấu.\n"
        else:
            comment += "- Không đủ dữ liệu để đánh giá tỷ lệ E/A.\n"

        return comment.strip()

    comment = generate_leverage_comment(ratio_df, stock_code)
    style = ParagraphStyle(
        name="GrowthComment",
        fontName="Roboto_Regular",
        fontSize=11,
        leading=15,
        alignment=TA_JUSTIFY,
    )
    para = Paragraph(comment.replace("\n", "<br/>"), style)
    w, h = para.wrap(width - 80, height)
    para.drawOn(c, 40, height - 660)

    # Ngắt sang trang mới
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Các chỉ số tài chính"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "CÁC CHỈ SỐ TÀI CHÍNH")

    # Tiêu đề "3. Chỉ số tăng trưởng (Growth Ratios)"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 127, "4. Chỉ số tăng trưởng (Growth Ratios)")

    c.setFillColor(colors.black)

    # Lấy dữ liệu từ industry_avg
    growth_company = ratio_df[(ratio_df['Mã'] == stock_code) & (ratio_df['Năm'] == 2024)]
    industry_row = industry_avg_df[
        (industry_avg_df['Ngành ICB - cấp 3'] == industry_name) & (industry_avg_df['Năm'] == 2024)
        ]

    if not growth_company.empty and not industry_row.empty:
        def extract(val):
            return round(val, 2) if pd.notna(val) else "NA"

        rev_growth_c = extract(growth_company["Revenue Growth (%)"].values[0])
        net_growth_c = extract(growth_company["Net Income Growth (%)"].values[0])

        rev_growth_i = extract(industry_row["Revenue Growth (%)"].values[0])
        net_growth_i = extract(industry_row["Net Income Growth (%)"].values[0])

        def compare_icon(val1, val2):
            if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                if val1 > val2:
                    return Image(os.path.join(base_path, "arrow_up_green.png"), width=10, height=10)
                elif val1 < val2:
                    return Image(os.path.join(base_path, "arrow_down_red.png"), width=10, height=10)
            return "-"

        data_growth = [
            ["Chỉ số", ten_cong_ty, "Trung bình ngành", "So sánh"],
            ["Revenue Growth (%)", rev_growth_c, rev_growth_i, compare_icon(rev_growth_c, rev_growth_i)],
            ["Net Income Growth (%)", net_growth_c, net_growth_i, compare_icon(net_growth_c, net_growth_i)],
        ]

        table_growth = Table(data_growth, colWidths=[150, 130, 130, 125])

        background_styles = [('BACKGROUND', (0, 0), (-1, 0), colors.white)]
        for i in range(1, len(data_growth)):
            bg_color = bg if i % 2 == 1 else colors.white
            background_styles.append(('BACKGROUND', (0, i), (-1, i), bg_color))

        table_growth.setStyle(TableStyle([
                                             ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                                             ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                                             ('ALIGN', (1, 1), (2, -1), 'LEFT'),
                                             ('ALIGN', (3, 1), (3, -1), 'LEFT'),

                                             ('FONTNAME', (0, 0), (-1, 0), 'Roboto_Black'),
                                             ('FONTNAME', (0, 1), (0, -1), 'Roboto_Regular'),
                                             ('FONTNAME', (1, 1), (2, -1), 'Roboto_Regular'),
                                             ('FONTSIZE', (0, 0), (-1, -1), 10),
                                             ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                                             ('TOPPADDING', (0, 0), (-1, -1), 6),
                                         ] + background_styles))

        # Vẽ bảng vào PDF
        table_growth.wrapOn(c, width, height)
        table_growth.drawOn(c, 40, height - 210)

    # Tiêu đề "Biểu đồ so sánh"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 230, "Biểu đồ so sánh")

    # Vẽ chart
    chart_buffer = draw_growth_chart(ratio_df, stock_code)
    if chart_buffer:
        img = Image(chart_buffer, width=520, height=300)
        img.wrapOn(c, width, height)
        img.drawOn(c, 40, height - 540)

    # Tiêu đề "Nhận xét"
    c.setFont("Roboto_Black", 14)
    c.drawString(60, height - 550, "Nhận xét")

    def generate_growth_comment(ratio_df, stock_code):
        df_plot = ratio_df[(ratio_df["Mã"] == stock_code) & (ratio_df["Năm"].between(2020, 2024))].sort_values("Năm")
        if df_plot.empty:
            return "Không đủ dữ liệu để đưa ra nhận xét cụ thể về xu hướng tăng trưởng doanh thu và lợi nhuận trong giai đoạn 2020–2024."

        rev_growth = df_plot["Revenue Growth (%)"].dropna()
        net_growth = df_plot["Net Income Growth (%)"].dropna()

        avg_rev = rev_growth.mean() if not rev_growth.empty else None
        avg_net = net_growth.mean() if not net_growth.empty else None

        comment = "Phân tích xu hướng tăng trưởng doanh thu và lợi nhuận sau thuế của công ty trong giai đoạn 2020–2024 cho thấy:\n\n"

        # Doanh thu
        if avg_rev is not None:
            if avg_rev > 10:
                comment += (
                    f"- Doanh thu có tốc độ tăng trưởng ấn tượng, với mức trung bình hàng năm đạt khoảng {avg_rev:.2f}%. "
                    "Điều này phản ánh khả năng mở rộng thị trường, phát triển sản phẩm hoặc dịch vụ mới hiệu quả của doanh nghiệp. "
                    "Một xu hướng như vậy thường là tín hiệu tích cực đối với các nhà đầu tư, bởi nó cho thấy công ty có nền tảng tăng trưởng bền vững trong dài hạn.\n\n"
                )
            elif avg_rev > 0:
                comment += (
                    f"- Doanh thu ghi nhận mức tăng trưởng trung bình {avg_rev:.2f}% mỗi năm. "
                    "Dù không thực sự bứt phá, đây vẫn là dấu hiệu cho thấy công ty duy trì được đà tăng trưởng ổn định, "
                    "mặc dù tốc độ này có thể chưa đủ mạnh để tạo ra lợi thế cạnh tranh rõ nét trên thị trường.\n\n"
                )
            else:
                comment += (
                    f"- Doanh thu có chiều hướng giảm nhẹ, với mức trung bình khoảng {avg_rev:.2f}%. "
                    "Sự sụt giảm này có thể phản ánh những khó khăn trong việc duy trì thị phần, hoặc ảnh hưởng từ yếu tố bên ngoài như điều kiện kinh tế vĩ mô. "
                    "Nếu xu hướng này tiếp tục kéo dài, công ty cần nhanh chóng đánh giá lại chiến lược kinh doanh để tránh rơi vào tình trạng suy giảm kéo dài.\n\n"
                )
        else:
            comment += "- Không có đủ dữ liệu để đánh giá xu hướng tăng trưởng doanh thu trong giai đoạn này.\n\n"

        # Lợi nhuận sau thuế
        if avg_net is not None:
            if avg_net > 10:
                comment += (
                    f"- Lợi nhuận sau thuế tăng trưởng mạnh mẽ, trung bình đạt khoảng {avg_net:.2f}% mỗi năm. "
                    "Đây là dấu hiệu rõ ràng cho thấy công ty không chỉ tăng doanh thu mà còn kiểm soát tốt chi phí, cải thiện hiệu quả hoạt động. "
                    "Tăng trưởng lợi nhuận như vậy góp phần củng cố niềm tin của nhà đầu tư vào triển vọng dài hạn của doanh nghiệp.\n"
                )
            elif avg_net > 0:
                comment += (
                    f"- Lợi nhuận sau thuế tăng trưởng với mức trung bình {avg_net:.2f}%. "
                    "Dù chưa thực sự bứt phá, nhưng vẫn cho thấy công ty đang đi đúng hướng trong việc nâng cao hiệu quả kinh doanh. "
                    "Tuy nhiên, công ty cần tiếp tục tối ưu hóa biên lợi nhuận để chuyển đổi tăng trưởng doanh thu thành lợi nhuận tốt hơn.\n"
                )
            else:
                comment += (
                    f"- Lợi nhuận sau thuế có dấu hiệu suy giảm, với mức trung bình {avg_net:.2f}%. "
                    "Đây có thể là hệ quả từ việc chi phí vận hành tăng nhanh hơn doanh thu, hoặc những yếu tố bất lợi như chi phí tài chính, thuế, hay biến động thị trường. "
                    "Việc lợi nhuận sụt giảm là tín hiệu cần được theo dõi chặt chẽ vì có thể ảnh hưởng đến khả năng sinh lời và phân phối cổ tức trong tương lai.\n"
                )
        else:
            comment += "- Không có đủ dữ liệu để đánh giá xu hướng tăng trưởng lợi nhuận sau thuế trong giai đoạn này.\n"

        return comment.strip()

    comment = generate_growth_comment(ratio_df, stock_code)

    style = ParagraphStyle(
        name="GrowthComment",
        fontName="Roboto_Regular",
        fontSize=11,
        leading=15,
        alignment=TA_JUSTIFY,
    )
    para = Paragraph(comment.replace("\n", "<br/>"), style)
    w, h = para.wrap(width - 80, height)
    para.drawOn(c, 40, height - 730)

    # Ngắt sang trang mới
    add_page_footer(c, width)
    c.showPage()

    # Header nền xám + tên công ty đỏ
    c.setFillColor(bg)
    c.rect(0, height - 90, width, 90, fill=True, stroke=0)

    c.setFont("Roboto_Black", 16)
    c.setFillColor(red_text)
    c.drawRightString(width - 40, height - 40, ten_cong_ty)

    c.setFont("Roboto_Black", 12)
    c.setFillColor(colors.black)
    c.drawRightString(width - 40, height - 60, f"Giá đóng cửa: {int(stock_price):,} VND")
    c.drawRightString(width - 40, height - 80, f"Ngày báo cáo: {ngay_tao}")

    # Line
    c.setStrokeColor(grey)
    c.setLineWidth(1.5)
    c.line(40, height - 90, width - 40, height - 90)
    # Tiêu đề "Khuến nghị dành cho nhà đầu tư"
    c.setFont("Roboto_Black", 14)
    c.setFillColor(red_text)
    c.drawString(40, height - 107, "KHUYẾN NGHỊ DÀNH CHO NHÀ ĐẦU TƯ")

    c.setFillColor(colors.black)

    def generate_investment_recommendation(
            roe_c, roe_i, roa_c, roa_i, ros_c, ros_i,
            pe_c, pe_i, pb_c, pb_i,
            cr_c, cr_i, qr_c, qr_i,
            rev_growth, net_growth,
            de_c, de_i
    ):
        comment = "Đánh giá tổng hợp và khuyến nghị đầu tư:\n\n"

        # 1. Khả năng sinh lời
        if all(isinstance(x, (int, float)) for x in [roe_c, roe_i, roa_c, roa_i, ros_c, ros_i]):
            if roe_c > roe_i and roa_c > roa_i and ros_c > ros_i:
                comment += (
                    "- Công ty đang có hiệu suất sinh lời vượt trội so với trung bình ngành. "
                    "Điều này cho thấy năng lực quản lý hiệu quả và mô hình kinh doanh đang hoạt động tốt.\n"
                )
            elif roe_c < roe_i and roa_c < roa_i and ros_c < ros_i:
                comment += (
                    "- Các chỉ số sinh lời đều thấp hơn ngành, phản ánh hiệu quả hoạt động còn hạn chế. "
                    "Nhà đầu tư nên cẩn trọng và theo dõi khả năng cải thiện trong tương lai.\n"
                )
            else:
                comment += (
                    "- Hiệu suất sinh lời có sự chênh lệch giữa các chỉ số, phản ánh chiến lược hoạt động "
                    "và cấu trúc tài chính có điểm khác biệt so với ngành.\n"
                )

        # 2. Định giá
        if all(isinstance(x, (int, float)) for x in [pe_c, pe_i, pb_c, pb_i]):
            if pe_c < pe_i and pb_c < pb_i:
                comment += (
                    "- Cổ phiếu có mức định giá thấp hơn trung bình ngành. "
                    "Đây có thể là cơ hội đầu tư tiềm năng nếu các yếu tố cơ bản được duy trì ổn định.\n"
                )
            elif pe_c > pe_i and pb_c > pb_i:
                comment += (
                    "- Cổ phiếu đang được thị trường định giá cao hơn ngành. "
                    "Điều này phản ánh kỳ vọng tích cực nhưng cũng tiềm ẩn rủi ro nếu kỳ vọng không được hiện thực hóa.\n"
                )
            else:
                comment += (
                    "- Định giá của công ty không quá chênh lệch so với ngành, "
                    "nhà đầu tư nên đánh giá thêm yếu tố tăng trưởng và dòng tiền.\n"
                )

        # 3. Khả năng thanh khoản
        if all(isinstance(x, (int, float)) for x in [cr_c, cr_i, qr_c, qr_i]):
            if cr_c > cr_i and qr_c > qr_i:
                comment += (
                    "- Công ty có khả năng thanh toán ngắn hạn tốt hơn trung bình ngành, "
                    "phản ánh việc quản lý tài sản lưu động hiệu quả và an toàn tài chính tốt.\n"
                )
            elif cr_c < cr_i and qr_c < qr_i:
                comment += (
                    "- Chỉ số thanh khoản của công ty thấp hơn ngành, cần theo dõi kỹ dòng tiền hoạt động "
                    "và các khoản phải thu.\n"
                )

        # 4. Tăng trưởng
        if rev_growth > 10 and net_growth > 10:
            comment += (
                "- Công ty có tốc độ tăng trưởng doanh thu và lợi nhuận cao, "
                "cho thấy tiềm năng mở rộng kinh doanh và hiệu quả vận hành tốt.\n"
            )
        elif rev_growth < 0 or net_growth < 0:
            comment += (
                "- Tăng trưởng đang chậm lại hoặc suy giảm. "
                "Nhà đầu tư nên theo dõi các chiến lược phục hồi và tái cấu trúc (nếu có).\n"
            )

        # 5. Đòn bẩy tài chính
        if isinstance(de_c, (int, float)) and isinstance(de_i, (int, float)):
            if de_c > de_i:
                comment += (
                    "- Tỷ lệ nợ cao hơn trung bình ngành, làm tăng rủi ro tài chính "
                    "trong bối cảnh thị trường biến động.\n"
                )
            else:
                comment += (
                    "- Công ty đang kiểm soát đòn bẩy tài chính hợp lý, giúp đảm bảo tính ổn định trong vận hành.\n"
                )

        # 6. Kết luận chung
        comment += "\nKết luận:\n"
        if roe_c > roe_i and pe_c < pe_i and rev_growth > 5 and de_c <= de_i:
            comment += (
                "--> Tổng thể, cổ phiếu này thể hiện tiềm năng tốt với hiệu suất sinh lời cao, "
                "định giá hợp lý và cơ cấu tài chính an toàn. Có thể xem xét đầu tư trong trung – dài hạn."
            )
        elif roe_c < roe_i and rev_growth < 0:
            comment += (
                "--> Cổ phiếu chưa nổi bật so với trung bình ngành về hiệu quả và tăng trưởng. "
                "Nhà đầu tư nên theo dõi thêm các tín hiệu cải thiện trước khi giải ngân."
            )
        else:
            comment += (
                "--> Cổ phiếu có điểm mạnh nhất định nhưng vẫn tồn tại yếu tố cần theo dõi. "
                "Phù hợp với nhà đầu tư có khẩu vị rủi ro trung bình hoặc dài hạn."
            )

        return comment.strip()

    def extract(val):
        return round(val, 2) if pd.notna(val) else "NA"

    pe_c = extract(company_row["P/E"].values[0])
    pe_i = extract(industry_row["P/E"].values[0])
    pb_c = extract(company_row["P/B"].values[0])
    pb_i = extract(industry_row["P/B"].values[0])

    cr_c = extract(company_row["Current Ratio"].values[0])
    cr_i = extract(industry_row["Current Ratio"].values[0])
    qr_c = extract(company_row["Quick Ratio"].values[0])
    qr_i = extract(industry_row["Quick Ratio"].values[0])

    avg_rev_growth = ratio_df.loc[
        (ratio_df["Mã"] == stock_code) & (ratio_df["Revenue Growth (%)"].notna()), "Revenue Growth (%)"
    ].mean()

    avg_net_growth = ratio_df.loc[
        (ratio_df["Mã"] == stock_code) & (ratio_df["Net Income Growth (%)"].notna()), "Net Income Growth (%)"
    ].mean()

    recommend_text = generate_investment_recommendation(
        roe_c, roe_i, roa_c, roa_i, ros_c, ros_i,
        pe_c, pe_i, pb_c, pb_i,
        cr_c, cr_i, qr_c, qr_i,
        avg_rev_growth, avg_net_growth,
        de_c, de_i
    )
    
    # Dùng thẻ <br/> thay vì \n để ngắt dòng trong Paragraph
    recommend_text_html = recommend_text.replace("\n", "<br/>")

    recommend_paragraph = Paragraph(recommend_text_html, styleN)
    w, h_recommend = recommend_paragraph.wrap(width - 80, height)
    recommend_paragraph.drawOn(c, 40, height - 115 - h_recommend)

    add_page_footer(c, width)

    c.save()
    buffer.seek(0)
    return buffer, ten_cong_ty

# Giao diện Streamlit
st.title("📊 XEM THÔNG TIN CHỨNG KHOÁN")
st.write("Nhấn nút Tạo PDF để tạo, sau đó ấn tiếp Tải PDF để xem")

stock_code = st.selectbox("Chọn mã chứng khoán", df["Mã"].unique())
selected_date = st.date_input("Chọn ngày báo cáo", min_value=min_date, max_value=max_date, value=max_date)

if st.button("📥 Tạo PDF"):
    pdf_buffer, ten_cong_ty = create_pdf(stock_code, selected_date)
    if pdf_buffer:
        file_name = f"Thông tin doanh nghiệp {ten_cong_ty}.pdf"
        st.download_button(label="📥 Tải PDF", data=pdf_buffer, file_name=file_name, mime="application/pdf")
    else:
        st.error("Không tìm thấy thông tin mã chứng khoán!")