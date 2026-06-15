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

# --- LINE Pay 票券新規大量上架優化版主程式 ---
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="LINE Pay 新規上架工具", layout="wide")

if "clear_id" not in st.session_state:
    st.session_state["clear_id"] = 0

cid = st.session_state["clear_id"]

st.title("🚀 LINE Pay 票券大量商品上架標準 Excel 生成器 (新規優化版)")
st.error("💡 官方規格提醒：不同規格的商品需要分開上架（例如：大、中、小杯，或是冷、熱飲，皆須拆分為獨立的商品品項填寫）。")

st.write("---")

# 1. 商店與品牌基本資料
st.subheader("🏢 1. 實際商品(服務)提供者基本資料填寫")
c1, c2, c3 = st.columns(3)
with c1:
    brand_name = st.text_input("品牌名稱 (兼商店名稱)", placeholder="例如：笨道策展咖啡", key=f"brand_name_{cid}")
    store_tax_id = st.text_input("商店統編", value="24941093", key=f"store_tax_id_{cid}")
with c2:
    store_address = st.text_input("商店地址", value="台北市南港區", key=f"store_address_{cid}")
    store_phone1 = st.text_input("商店電話 1", value="0912345678", key=f"store_phone1_{cid}")
with c3:
    store_phone2 = st.text_input("商店電話 2", value="02-6631-5166", key=f"store_phone2_{cid}")

st.write("---")

# 2. 三大公版條款文字庫定義 (全新修正版)
catering_default = (
    "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n"
    "(2)點餐前，請事先告知服務人員欲使用本券，結帳時出示本券畫面予櫃台掃碼使用。\n"
    "(3)飲料可加價升級特大杯或添加客製化等收費項目。\n"
    "(4)兌換數量依各門市現貨為準，若該門市已無存貨，請至鄰近門市兌換，或指定飲料可選擇更換其他系列品項。\n"
    "(5)本券平假日皆適用。\n"
    "(6)本券僅限於台灣區域使用且不適用於外送服務與網路訂餐。\n"
    "(7)本券商品不得與其他優惠合併使用且一桌限用一張。\n"
    "(8)本券商品無法累積會員積分。\n"
    "(9)本商品恕無法折抵商場停車。"
)

beauty_default = (
    "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n"
    "(2)本券平假日皆適用。\n"
    "(3)本券商品無法累積會員積分。\n"
    "(4)本券不得指定美容師、療程房型；如您是敏感膚質，或有生理期、懷孕、心血管疾病等特殊身體狀況，請於課程開始前，主動告知美容師。\n"
    "(5)兌換數量依各門市現貨為準，若該門市已無存貨，請至鄰近門市兌換，或指定飲料可選擇更換其他系列品項。"
)

entertainment_default = (
    "(1)本優惠為單人使用，不建議 12 歲以下孩童參與。\n"
    "(2)本遊戲 4 人成團，未滿 8 人需與其他玩家併團遊戲。\n"
    "(3)本優惠活動時間約 100 分鐘(含講解)。"
)

# 3. 商品動態輸入區 (1 ~ 5 項商品)
st.subheader("📦 2. 商品品項與規格細節填寫")

products_data = []

for i in range(1, 6):
    with st.expander(f"🔹 點我填寫第 {i} 項商品內容", expanded=(i == 1)):
        p_c1, p_c2 = st.columns(2)
        
        with p_c1:
            p_code = st.text_input(f"商品序號 ({i})", placeholder="例如：001", key=f"p_code_{i}_{cid}")
            p_name = st.text_input(f"商品名稱/規格 ({i})", placeholder="例如：美式咖啡(大杯/冷)", key=f"p_name_{i}_{cid}")
            p_orig_price = st.text_input(f"商品原價 ({i})", placeholder="例如：150", key=f"p_orig_price_{i}_{cid}")
            p_disc_price = st.text_input(f"商品優惠價 ({i})", placeholder="例如：120", key=f"p_disc_price_{i}_{cid}")
            p_stock = st.text_input(f"庫存數 ({i})", placeholder="例如：1000", key=f"p_stock_{i}_{cid}")
            p_desc = st.text_area(f"商品描述及文案說明 ({i})", placeholder="關於商品的詳細說明（字數限制為 32,767 個字元）", key=f"p_desc_{i}_{cid}", height=100)

        with p_c2:
            # 📅 商品券兌換時間
            st.write("**📅 商品券兌換時間設定**")
            validity_type = st.radio(f"兌換類型 ({i})", ["常態商品 (系統自訂50天)", "季節商品 (需滿足 截止日 - 初日 ≤ 50天)"], key=f"v_type_{i}_{cid}")
            
            t_now = datetime.now()
            if validity_type == "常態商品 (系統自訂50天)":
                start_d = t_now.strftime("%Y/%m/%d")
                end_d = (t_now + timedelta(days=50)).strftime("%Y/%m/%d")
                st.info(f"常態商品自動帶入區間：{start_d} 至 {end_d}")
            else:
                vd1, vd2 = st.columns(2)
                with vd1:
                    start_date_val = st.date_input(f"上架販售初日 ({i})", value=t_now, key=f"s_date_{i}_{cid}")
                with vd2:
                    end_date_val = st.date_input(f"兌換最後截止日 ({i})", value=t_now + timedelta(days=45), key=f"e_date_{i}_{cid}")
                
                # 計算是否超過50天限制
                days_diff = (end_date_val - start_date_val).days
                if days_diff > 50:
                    st.error(f"⚠️ 警告：目前相差 {days_diff} 天，違反 LINE Pay 官方「季節商品需 ≤ 50天」的規定！")
                else:
                    st.success(f"✅ 符合規定（目前相差 {days_diff} 天）")
                    
                start_d = start_date_val.strftime("%Y/%m/%d")
                end_d = end_date_val.strftime("%Y/%m/%d")
            
            # 文案格式選擇
            format_type = st.selectbox(
                f"選擇要帶入的官方兌換時間文案格式 ({i})",
                [
                    "格式一：• 本券兌換期間為（YYYY/MM/DD 至 YYYY/MM/DD）",
                    "格式二：• 本券可兌換（商品名稱/規格），兌換期間為購買當日起至（兌換天數50日）日止。",
                    "格式三：• 本券可兌換（商品名稱/規格），兌換期間為（YYYY/MM/DD 至 YYYY/MM/DD）"
                ],
                key=f"format_type_{i}_{cid}"
            )
            
            display_name = p_name if p_name else f"（第{i}項商品名稱/規格）"
            if "格式一" in format_type:
                final_validity_text = f"• 本券兌換期間為（{start_d} 至 {end_d}）"
            elif "格式二" in format_type:
                final_validity_text = f"• 本券可兌換（{display_name}），兌換期間為購買當日起至（兌換天數50日）日止。"
            else:
                final_validity_text = f"• 本券可兌換（{display_name}），兌換期間為（{start_d} 至 {end_d}）"
            st.code(final_validity_text, language="text")

            # 🛒 商品販售時間
            st.write("**🛒 商品販售時間**")
            sell_type = st.radio(f"販售屬性 ({i})", ["常態商品 (填無)", "季節性商品 (填日期區間)"], key=f"sell_type_{i}_{cid}")
            if sell_type == "常態商品 (填無)":
                final_sell_time = "無"
            else:
                s1, s2 = st.columns(2)
                with s1:
                    s_s = st.date_input(f"販售開始日 ({i})", value=t_now, key=f"s_s_{i}_{cid}")
                with s2:
                    s_e = st.date_input(f"販售結束日 ({i})", value=t_now + timedelta(days=14), key=f"s_e_{i}_{cid}")
                final_sell_time = f"{s_s.strftime('%Y/%m/%d')} 至 {s_e.strftime('%Y/%m/%d')}"

            # 🏷️ 商品折扣時間
            st.write("**🏷️ 商品折扣時間**")
            discount_type = st.radio(f"折扣屬性 ({i})", ["原價販售 (填無)", "限定折扣 (填日期區間)"], key=f"discount_type_{i}_{cid}")
            if discount_type == "原價販售 (填無)":
                final_discount_time = "無"
            else:
                d1, d2 = st.columns(2)
                with d1:
                    d_s = st.date_input(f"折扣開始日 ({i})", value=t_now, key=f"d_s_{i}_{cid}")
                with d2:
                    d_e = st.date_input(f"折扣結束日 ({i})", value=t_now + timedelta(days=7), key=f"d_e_{i}_{cid}")
                final_discount_time = f"{d_s.strftime('%Y/%m/%d')} 至 {d_e.strftime('%Y/%m/%d')}"

        st.write("---")
        # 📜 使用條款下拉選單
        st.write("**📜 使用條款說明公版選單**")
        terms_template = st.selectbox(
            f"選擇適用之產業條款範本 ({i})",
            ["餐飲公版條款", "美容公版條款", "休閒娛樂公版條款", "完全自訂空白"],
            key=f"terms_template_{i}_{cid}"
        )
        
        if terms_template == "餐飲公版條款":
            terms_value = catering_default
        elif terms_template == "美容公版條款":
            terms_value = beauty_default
        elif terms_template == "休閒娛樂公版條款":
            terms_value = entertainment_default
        else:
            terms_value = ""
            
        p_terms = st.text_area(f"編輯/修正使用條款內容 ({i})", value=terms_value, key=f"p_terms_{i}_{terms_template}_{cid}", height=150)
        p_main_img = st.file_uploader(f"上傳商品圖片 ({i})（系統將自動縮放規格為 W640 x H640 px）", type=["jpg", "jpeg", "png"], key=f"p_main_img_{i}_{cid}")

        if p_name:
            products_data.append({
                "code": p_code,
                "name": p_name,
                "orig_price": p_orig_price,
                "disc_price": p_disc_price,
                "stock": p_stock,
                "validity": final_validity_text,
                "desc": p_desc,
                "terms": p_terms,
                "sell_time": final_sell_time,
                "discount_time": final_discount_time,
                "main_img": p_main_img
            })

st.write("---")

# 4. 底部按鈕區
col_btn1, col_btn2 = st.columns([1, 4])

with col_btn1:
    if st.button("🧹 清除所有內容（做下一家店）", use_container_width=True):
        if "final_excel_bytes" in st.session_state:
            del st.session_state["final_excel_bytes"]
        if "final_file_name" in st.session_state:
            del st.session_state["final_file_name"]
        st.session_state["clear_id"] += 1
        st.rerun()

with col_btn2:
    generate_pressed = st.button("🚀 生成完整票券 Excel 資料表", type="primary", use_container_width=True)

if generate_pressed:
    if len(products_data) == 0:
        st.error("❌ 請至少填寫一項商品的『商品名稱/規格』才能進行 Excel 匯出！")
    else:
        with st.spinner("⏳ 正在為您排版並將所有上傳的照片高精細裁切為 W640xH640，請稍候..."):
            
            # 固定不可變動法規文字組裝
            b_display = brand_name if brand_name else "（請填寫品牌名稱）"
            notices = (
                f"1. 使用本券請至 {b_display} 直接出示本券掃碼兌換（請將螢幕亮度調到最大）。\n"
                "2. 本券恕不得更換現金及轉售。\n"
                "3. 使用本券時須符合本券載明之品項與規格。因購買時LINE Pay已開立發票給購買者，兌換時不另開立發票。商品兌換後，恕無法提供退貨及換貨。\n"
                "4. 商店僅提供兌換本券商品的服務，若對兌換之商品有任何問題請洽門市人員，其他本服務相關問題請聯繫連加網路商業股份有限公司（下稱LINE Pay）客服。\n"
                "5. 本券不記名，僅限兌換一次，不得重複使用，任何人持有皆可兌換，請自行妥善保管。\n"
                "6. 本券之兌換與銷售，恕不與商店所有折扣、優惠、各行銷活動合併使用。\n"
                "7. 有關本券之使用、兌換、取消及補發之條款及條件，及本服務之完整內容，請詳見「服務條款」。\n"
                "8. 本券如未於期限內兌換，費用將全額退款給原購買者。"
            )
            
            issuer_info = "連加網路商業股份有限公司\n地址：臺北市南港區經貿二路121號18樓\n電話：02-3518-7600\n統編：24941093"
            guarantee_info = "本服務所發行之票券金額，皆自發行日起存入發行人於國泰世華商業銀行開立之信託帳戶，專款專用。所謂專用，係指供發行人履行交付商品或提供服務義務使用，前述信託期間自出售日起算至少一年。"
            provider_info = f"商店名稱：{b_display}\n地址：{store_address}\n電話：{store_phone1} / {store_phone2}\n統編：{store_tax_id}"

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "LINEPay新規上架資料"
            ws.views.sheetView[0].showGridLines = True
            
            headers = [
                "商品序號", "品牌名稱", "商品名稱/規格", "商品圖片(640x640)", "商品券兌換時間說明",
                "商品描述及文案說明", "使用條款說明", "商品販售時間", "原價", "優惠價",
                "商品折扣時間", "庫存數", "禮券發行者(固定)", "履約保證(固定)", "注意事項說明", "實際商品(服務)提供者"
            ]
            ws.append(headers)
            
            # 高級商務深藍配配色排版
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
            
            col_widths = {
                'A': 15, 'B': 18, 'C': 30, 'D': 45, 'E': 45, 'F': 45, 'G': 55,
                'H': 30, 'I': 12, 'J': 12, 'K': 30, 'L': 12, 'M': 35, 'N': 45, 'O': 55, 'P': 40
            }
            for col_let, width in col_widths.items():
                ws.column_dimensions[col_let].width = width
            
            ws.row_dimensions[1].height = 35
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = title_font
                cell.fill = title_fill
                cell.alignment = title_alignment
                cell.border = thin_border

            # 影像記憶體安全保溫箱，徹底杜絕存檔崩潰 Bug
            img_keep_alive = []

            current_row = 2
            for p in products_data:
                ws.row_dimensions[current_row].height = 250
                
                ws.cell(row=current_row, column=1, value=p["code"])
                ws.cell(row=current_row, column=2, value=brand_name)
                ws.cell(row=current_row, column=3, value=p["name"])
                ws.cell(row=current_row, column=5, value=p["validity"])
                ws.cell(row=current_row, column=6, value=p["desc"])
                ws.cell(row=current_row, column=7, value=p["terms"])
                ws.cell(row=current_row, column=8, value=p["sell_time"])
                ws.cell(row=current_row, column=9, value=p["orig_price"])
                ws.cell(row=current_row, column=10, value=p["disc_price"])
                ws.cell(row=current_row, column=11, value=p["discount_time"])
                ws.cell(row=current_row, column=12, value=p["stock"])
                ws.cell(row=current_row, column=13, value=issuer_info)
                ws.cell(row=current_row, column=14, value=guarantee_info)
                ws.cell(row=current_row, column=15, value=notices)
                ws.cell(row=current_row, column=16, value=provider_info)
                
                # 自動剪切轉為官方標準 W640 x H640 px PNG 圖層
                if p["main_img"]:
                    pm_img = Image.open(p["main_img"]).resize((640, 640))
                    pm_buf = io.BytesIO()
                    pm_img.save(pm_buf, format='PNG')
                    pm_buf.seek(0)
                    img_keep_alive.append(pm_buf)
                    ws.add_image(OpenpyxlImage(pm_buf), f"D{current_row}")

                for col_idx in range(1, len(headers) + 1):
                    cell = ws.cell(row=current_row, column=col_idx)
                    cell.font = data_font
                    cell.border = thin_border
                    if col_idx in [1, 2, 5, 8, 9, 10, 11, 12]:
                        cell.alignment = center_align
                    else:
                        cell.alignment = text_align
                
                current_row += 1

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            
            st.session_state["final_excel_bytes"] = excel_buffer.getvalue()
            st.session_state["final_file_name"] = f"{brand_name if brand_name else 'LINEPay'}_新規票券規格資料表.xlsx"

if "final_excel_bytes" in st.session_state:
    st.write("")
    st.success("🎉 格式完全優化成功！兌換文案與產業選單已完美嵌入 Excel。請點擊下方下載：")
    st.download_button(
        label="💾 點我立刻下載 2026 新規上架 Excel 檔案",
        data=st.session_state["final_excel_bytes"],
        file_name=st.session_state["final_file_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
