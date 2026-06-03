import os
import sys
import subprocess
import importlib

# 🚀 終極防錯：在網頁資料夾內建立專屬工具箱，強制繞過雲端系統安裝 Bug
local_libs = os.path.join(os.getcwd(), "local_packages")
if local_libs not in sys.path:
    sys.path.insert(0, local_libs)

for package, import_name in [("openpyxl", "openpyxl"), ("Pillow", "PIL")]:
    try:
        importlib.import_module(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--target", local_libs, package])
        importlib.invalidate_caches()

# --- LINE Pay 票券專用主程式 ---
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="LINE Pay 票券工具", layout="wide")
st.title("📦 LINE Pay 票券商品資料完整轉 Excel 工具")
st.write("請在下方填寫商店與商品資料，系統將自動套用定型化契約條款並生成標準 Excel。")

# 使用 Streamlit 區塊區分填寫內容
col1, col2 = st.columns(2)

with col1:
    st.subheader("🏢 商店與品牌基本資料")
    brand_name = st.text_input("品牌名稱", placeholder="例如：星巴克 / 必勝客")
    store_name = st.text_input("實際提供者 - 商店名稱", placeholder="例如：台北信義店")
    store_address = st.text_input("實際提供者 - 地址")
    store_phone = st.text_input("實際提供者 - 電話")

with col2:
    st.subheader("🛍️ 商品與價格資料")
    prod_name = st.text_input("本券可兌換（商品名稱）", placeholder="例如：大杯那堤 / 雙層美式綜合比薩")
    original_price = st.text_input("原價（數字）", placeholder="例如：150")
    discount_price = st.text_input("優惠價（數字）", placeholder="例如：120")
    
    uploaded_file = st.file_uploader("商品主圖（會自動調整為 375x375 規格）", type=["jpg", "jpeg", "png"])

st.subheader("📝 商品詳情說明")
description = st.text_area("商品描述", placeholder="填寫商品的特色、口感或服務介紹...", height=100)
specifications = st.text_area("商品規格", placeholder="例如：容量：473ml / 尺寸：9吋...", height=100)

# 點擊生成按鈕
if st.button("🚀 生成完整版票券 Excel"):
    if not brand_name or not prod_name:
        st.error("❌ 請務必填寫「品牌名稱」與「商品名稱」！")
    else:
        # 1. 自動計算日期（今天 至 今天 + 60 天）
        today_str = datetime.now().strftime("%Y-%m-%d")
        future_str = (datetime.now() + timedelta(days=60)).strftime("%Y-%m-%d")
        coupon_validity = f"{today_str} 至 {future_str} (上架後60天)"
        
        # 2. 組合固定條款文字
        terms_of_use = (
            "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n"
            "(2)點餐前，請事先告知服務人員欲使用本券，結帳時出示本券畫面予櫃台掃碼使用。\n"
            "(3)飲料可加價升級特大杯或添加客製化等收費項目。\n"
            "(4)兌換數量依各門市現貨為準，若該門市已無存貨，可選擇更換其他系列品項\n"
            "(5)本券平假日皆適用。\n"
            "(6)本券僅限於台灣區域使用且不適用於外送服務與網路訂餐。\n"
            "(7)本券商品不得與其他優惠合併使用。"
        )
        
        notices = (
            f"1. 使用本券請至 {brand_name} 直接出示本券掃碼兌換（請將螢幕亮度調到最大）。\n"
            "2. 本券恕不得更換現金及轉售。\n"
            "3. 使用本券時須符合本券載明之品項與規格。因購買時LINE Pay已開立發票給購買者，兌換時不另開立發票。商品兌換後，恕無法提供退貨及換貨。\n"
            "4. 商店僅提供兌換本券商品的服務，若對兌換之商品有任何問題請洽門市人員，其他本服務相關問題請聯繫連加網路商業股份有限公司（下稱LINE Pay）客服。\n"
            "5. 本券不記名，僅限兌換一次，不得重複使用，任何人持有皆可兌換，請自行妥善保管。\n"
            "6. 本券之兌換與銷售，恕不與商店所有折扣、優惠、各行銷活動合併使用。\n"
            "7. 有關本券之使用、兌換、取消及補發之條款及條件，及本服務之完整內容，請詳見「服務條款」。\n"
            "8. 本券如未於期限內兌換，費用將全額退款給原購買者。"
        )
        
        provider_info = f"商店名稱：{store_name}\n地址：{store_address}\n電話：{store_phone}"
        
        issuer_info = (
            "連加網路商業股份有限公司\n"
            "地址：台北市中山區敬業一路2號13樓\n"
            "電話：02-6631-5166\n"
            "統編：24941093"
        )
        
        guarantee_info = "本服務所發行之票券金額，皆自發行日起存入發行人於國泰世華商業銀行開立之信託帳戶，專款專用。所謂專用，係指供發行人履行交付商品或提供服務義務使用，前述信託期間自出售日起算至少一年。"
        
        # 3. 建立 Excel 工作簿
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "LINE Pay 票券上架資料"
        ws.views.sheetView[0].showGridLines = True # 強制顯示網格線
        
        # 4. 定義所有欄位 Headers (共 13 欄)
        headers = [
            "品牌名稱", "商品圖片", "原價", "優惠價", "兌換期", 
            "商品描述", "商品規格", "本券可兌換", "使用條款說明", 
            "注意事項", "實際商品(服務)提供者", "禮券發行者", "履約保證"
        ]
        ws.append(headers)
        
        # 5. 寫入填寫的資料
        ws["A2"] = brand_name
        ws["C2"] = original_price
        ws["D2"] = discount_price
        ws["E2"] = coupon_validity
        ws["F2"] = description
        ws["G2"] = specifications
        ws["H2"] = prod_name
        ws["I2"] = terms_of_use
        ws["J2"] = notices
        ws["K2"] = provider_info
        ws["L2"] = issuer_info
        ws["M2"] = guarantee_info
        
        # 6. 🎨 視覺美化排版
        ws.row_dimensions[1].height = 35 # 標題列高度
        
        # 精準設定每一欄的寬度，長文字欄位給予足夠空間
        column_widths = {
            'A': 18, 'B': 52, 'C': 12, 'D': 12, 'E': 30, 
            'F': 35, 'G': 35, 'H': 25, 'I': 55, 'J': 65, 
            'K': 35, 'L': 35, 'M': 50
        }
        for col_letter, width in column_widths.items():
            ws.column_dimensions[col_letter].width = width
            
        # 樣式定義
        title_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
        title_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid") # LINE Pay 質感深藍
        title_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        data_font = Font(name="微軟正黑體", size=10, color="000000")
        general_alignment = Alignment(horizontal="left", vertical="top", wrap_text=True) # 靠左、靠上、自動換行
        center_alignment = Alignment(horizontal="center", vertical="top", wrap_text=True) # 置中、靠上
        
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        
        # 套用標題樣式 (第 1 列)
        for col_num in range(1, 14):
            cell = ws.cell(row=1, column=col_num)
            cell.font = title_font
            cell.fill = title_fill
            cell.alignment = title_alignment
            cell.border = thin_border
            
        # 套用資料列樣式 (第 2 列)
        for col_num in range(1, 14):
            cell = ws.cell(row=2, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            # 針對價格與時間等欄位做置中對齊，文字類別靠左對齊
            if col_num in [3, 4, 5]:
                cell.alignment = center_alignment
            else:
                cell.alignment = general_alignment
                
        # 7. 📸 圖片縮放處理
        if uploaded_file is not None:
            img = Image.open(uploaded_file)
            img = img.resize((375, 375))
            
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            
            xl_img = OpenpyxlImage(img_byte_arr)
            ws.row_dimensions[2].height = 285 # 調整高度容納 375px 圖片
            ws.add_image(xl_img, 'B2')
        else:
            ws.row_dimensions[2].height = 150 # 若沒圖片，給予舒服的文字預設高度
            
        # 8. 匯出下載
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.download_button(
            label="💾 點我下載完整版 LINE Pay 上架 Excel",
            data=excel_buffer,
            file_name=f"{brand_name}_{prod_name}_完整上架資料.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("🎉 Excel 完整票券資料產生成功！")
