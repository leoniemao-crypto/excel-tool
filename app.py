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

# --- LINE Pay 票券終極大平台主程式 ---
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="LINE Pay 票券終極工具", layout="wide")
st.title("🚀 LINE Pay 票券大量商品上架標準 Excel 生成器")
st.write("本工具支援高達 5 項商品同時表單化輸入，並會自動將所有上傳的照片裁切、縮放至 LINE Pay 官方指定規格。")

st.write("---")

# 1. 商店與品牌基本資料
st.subheader("🏢 1. 商店與品牌基本資料（必填）")
c1, c2, c3 = st.columns(3)
with c1:
    mid = st.text_input("MID (商店代號)", placeholder="請輸入 LINE Pay MID", key="mid")
    brand_name = st.text_input("品牌名稱", placeholder="例如：笨道大飯店", key="brand_name")
with c2:
    store_name = st.text_input("實際提供者 - 商店名稱", placeholder="例如：台北總店", key="store_name")
    operating_hours = st.text_input("營業時間", placeholder="例如：週一至週日 11:00 - 22:00", key="operating_hours")
with c3:
    store_phone = st.text_input("實際提供者 - 電話", key="store_phone")
    store_address = st.text_input("實際提供者 - 地址", key="store_address")

# 2. 全域日期設定
default_validity = f"{(datetime.now()).strftime('%Y-%m-%d')} 至 {(datetime.now() + timedelta(days=60)).strftime('%Y-%m-%d')} (上架後60天)"
coupon_validity = st.text_input("兌換期設定（預設為上架後60天，可自行修正）", value=default_validity, key="coupon_validity")

st.write("---")

# 3. 店家全域照片上傳 (LOGO & Banner)
st.subheader("🖼️ 2. 店家形象照片上傳")
img_c1, img_c2 = st.columns(2)
with img_c1:
    logo_file = st.file_uploader("LOGO 照片（自動調整為 W300 x H300 px）", type=["jpg", "jpeg", "png"], key="logo_file")
with img_c2:
    banner_file = st.file_uploader("Banner 照片（自動調整為 W750 x H454 px）", type=["jpg", "jpeg", "png"], key="banner_file")

st.write("---")

# 4. 商品動態輸入區 (1 ~ 5 項商品)
st.subheader("📦 3. 商品品項填寫（最多可擴充至 5 項商品）")
st.caption("💡 提示：只要填寫了「商品名稱」，該品項就會被匯出至 Excel 中。")

products_data = []

for i in range(1, 6):
    with st.expander(f"🔹 點我填寫第 {i} 項商品內容", expanded=(i == 1)):
        p_c1, p_c2 = st.columns(2)
        with p_c1:
            p_name = st.text_input(f"本券可兌換（第 {i} 項商品名稱）", placeholder="例如：招牌牛肉麵 / 特大杯冷萃咖啡", key=f"p_name_{i}")
            p_orig_price = st.text_input(f"第 {i} 項商品原價", placeholder="例如：250", key=f"p_orig_price_{i}")
            p_disc_price = st.text_input(f"第 {i} 項商品優惠價", placeholder="例如：199", key=f"p_disc_price_{i}")
        with p_c2:
            p_main_img = st.file_uploader(f"第 {i} 項商品主圖（自動調整為 W640 x H640 px）", type=["jpg", "jpeg", "png"], key=f"p_main_img_{i}")
            p_other_imgs = st.file_uploader(f"第 {i} 項其它商品照片（最多10張，自動調整為 W640 x H640 px）", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key=f"p_other_imgs_{i}")
        
        p_desc = st.text_area(f"第 {i} 項商品描述 (上限 32,767 字)", placeholder="請輸入詳細商品介紹...", key=f"p_desc_{i}", height=80)
        p_spec = st.text_area(f"第 {i} 項商品規格", placeholder="例如：容量、尺寸、成份說明...", key=f"p_spec_{i}", height=80)
        
        # 如果有填名字，就塞入陣列準備處理
        if p_name:
            products_data.append({
                "index_str": f"商品{i}",
                "name": p_name,
                "orig_price": p_orig_price,
                "disc_price": p_disc_price,
                "main_img": p_main_img,
                "other_imgs": p_other_imgs[:10], # 強制最多10張
                "desc": p_desc,
                "spec": p_spec
            })

st.write("---")

# 5. 底部按鈕區（左：清除，右：生成）
col_btn1, col_btn2 = st.columns([1, 4])

with col_btn1:
    # 🧹 清除重新填寫，準備做下一家店
    if st.button("🧹 清除所有內容（做下一家店）", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

with col_btn2:
    # 🚀 生成完整 Excel
    if st.button("🚀 生成完整票券 Excel 資料表", type="primary", use_container_width=True):
        if not brand_name:
            st.error("❌ 請務必填寫『品牌名稱』才能生成定型化契約條款！")
        elif len(products_data) == 0:
            st.error("❌ 請至少填寫一項商品的『商品名稱』！")
        else:
            # 建立定型化公版條款
            terms_of_use = (
                "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n"
                "(2)點餐前，請事先告知服務人員欲使用本券，結帳時出示本券畫面予櫃台掃碼使用。\n"
                "(3)商品可加價升級或添加客製化等收費項目。\n"
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
            issuer_info = "連加網路商業股份有限公司\n地址：臺北市南港區經貿二路121號18樓\n電話：02-3518-7600\n統編：24941093"
            guarantee_info = "本服務所發行之票券金額，皆自發行日起存入發行人於國泰世華商業銀行開立之信託帳戶，專款專用。所謂專用，係指供發行人履行交付商品或提供服務義務使用，前述信託期間自出售日起算至少一年。"

            # 建立工作簿
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "LINE Pay 上架明細"
            ws.views.sheetView[0].showGridLines = True
            
            # 欄位 Headers 設計（共 30 欄）
            headers = [
                "MID", "品牌名稱", "營業時間", "兌換期", "實際商品提供者資訊", "禮券發行者", "履約保證",
                "LOGO照片(300x300)", "Banner照片(750x454)", "商品項次", "本券可兌換商品名稱", 
                "原價", "優惠價", "商品描述", "商品規格", "使用條款說明", "注意事項", 
                "商品主圖(640x640)", "其他照片1", "其他照片2", "其他照片3", "其他照片4", 
                "其他照片5", "其他照片6", "其他照片7", "其他照片8", "其他照片9", "其他照片10"
            ]
            ws.append(headers)
            
            # 美化樣式
            title_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
            title_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            title_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            data_font = Font(name="微軟正黑體", size=10)
            text_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
            center_align = Alignment(horizontal="center", vertical="top", wrap_text=True)
            thin_border = Border(
                left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
            )
            
            # 設定欄寬
            col_widths = {
                'A': 15, 'B': 18, 'C': 25, 'D': 30, 'E': 35, 'F': 35, 'G': 40,
                'H': 40, 'I': 45, 'J': 12, 'K': 25, 'L': 12, 'M': 12, 'N': 40, 'O': 35, 'P': 45, 'Q': 55,
                'R': 45, 'S': 15, 'T': 15, 'U': 15, 'V': 15, 'W': 15, 'X': 15, 'Y': 15, 'Z': 15, 'AA': 15, 'AB': 15
            }
            for col_let, width in col_widths.items():
                ws.column_dimensions[col_let].width = width
            
            # 格式化標題列
            ws.row_dimensions[1].height = 35
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = title_font
                cell.fill = title_fill
                cell.alignment = title_alignment
                cell.border = thin_border

            # 處理全域店照快取 (LOGO & Banner)
            logo_img_obj = None
            if logo_file:
                l_img = Image.open(logo_file).resize((300, 300))
                l_buf = io.BytesIO()
                l_img.save(l_buf, format='PNG')
                l_buf.seek(0)
                logo_img_obj = l_buf

            banner_img_obj = None
            if banner_file:
                b_img = Image.open(banner_file).resize((750, 454))
                b_buf = io.BytesIO()
                b_img.save(b_buf, format='PNG')
                b_buf.seek(0)
                banner_img_obj = b_buf

            # 循環寫入商品資料列
            current_row = 2
            for p in products_data:
                ws.row_dimensions[current_row].height = 250 # 固定高規格容納大圖
                
                # 寫入文字
                ws.cell(row=current_row, column=1, value=mid)
                ws.cell(row=current_row, column=2, value=brand_name)
                ws.cell(row=current_row, column=3, value=operating_hours)
                ws.cell(row=current_row, column=4, value=coupon_validity)
                ws.cell(row=current_row, column=5, value=provider_info)
                ws.cell(row=current_row, column=6, value=issuer_info)
                ws.cell(row=current_row, column=7, value=guarantee_info)
                
                ws.cell(row=current_row, column=10, value=p["index_str"])
                ws.cell(row=current_row, column=11, value=p["name"])
                ws.cell(row=current_row, column=12, value=p["orig_price"])
                ws.cell(row=current_row, column=13, value=p["disc_price"])
                ws.cell(row=current_row, column=14, value=p["desc"])
                ws.cell(row=current_row, column=15, value=p["spec"])
                ws.cell(row=current_row, column=16, value=terms_of_use)
                ws.cell(row=current_row, column=17, value=notices)
                
                # 塞入店照圖片
                if logo_img_obj:
                    logo_img_obj.seek(0)
                    ws.add_image(OpenpyxlImage(logo_img_obj), f"H{current_row}")
                if banner_img_obj:
                    banner_img_obj.seek(0)
                    ws.add_image(OpenpyxlImage(banner_img_obj), f"I{current_row}")
                
                # 處理商品主圖 (640x640)
                if p["main_img"]:
                    pm_img = Image.open(p["main_img"]).resize((640, 640))
                    pm_buf = io.BytesIO()
                    pm_img.save(pm_buf, format='PNG')
                    pm_buf.seek(0)
                    ws.add_image(OpenpyxlImage(pm_buf), f"R{current_row}")
                    
                # 處理其它
