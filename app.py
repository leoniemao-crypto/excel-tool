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

# --- LINE Pay 票券雙效終極大平台主程式 ---
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from PIL import Image
from datetime import datetime, timedelta
import io
import zipfile  # 內建庫，用於多張商品圖打包下載

st.set_page_config(page_title="LINE Pay 新規大平台", layout="wide")

if "clear_id" not in st.session_state:
    st.session_state["clear_id"] = 0

cid = st.session_state["clear_id"]

# 建立兩大服務分頁
tab1, tab2 = st.tabs(["📝 1. 票券純文字 Excel 生成", "🖼️ 2. 官方圖片尺寸調整工具"])

# ==========================================
# 📝 分頁一：純文字 Excel 生成邏輯
# ==========================================
with tab1:
    st.title("🚀 LINE Pay 票券大量商品上架標準 Excel 生成器")
    st.error("💡 官方規格提醒：不同規格的商品需要分開上架（例如：大、中、小杯，或是冷、熱飲，皆須拆分為獨立的商品品項填寫）。")
    
    st.write("---")
    st.subheader("🏢 實際商品(服務)提供者基本資料填寫")
    c1, c2, c3 = st.columns(3)
    with c1:
        brand_name = st.text_input("品牌名稱 (兼商店名稱)", placeholder="例如：笨道策展咖啡", key=f"brand_name_{cid}")
        store_tax_id = st.text_input("商店統編", value="24941093", key=f"store_tax_id_{cid}") # 💡 已修復：統一變數名稱
    with c2:
        store_address = st.text_input("商店地址", value="台北市南港區", key=f"store_address_{cid}") # 💡 已修復：統一變數名稱
        store_phone1 = st.text_input("商店電話 1", value="0912345678", key=f"store_phone1_{cid}") # 💡 已修復：統一變數名稱
    with c3:
        store_phone2 = st.text_input("商店電話 2", value="02-6631-5166", key=f"store_phone2_{cid}") # 💡 已修復：統一變數名稱

    # 三大公版條款文字庫
    catering_default = "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n(2)點餐前，請事先告知服務人員欲使用本券，結帳時出示本券畫面予櫃台掃碼使用。\n(3)飲料可加價升級特大杯或添加客製化等收費項目。\n(4)兌換數量依各門市現貨為準，若該門市已無存貨，請至鄰近門市兌換，或指定飲料可選擇更換其他系列品項。\n(5)本券平假日皆適用。\n(6)本券僅限於台灣區域使用且不適用於外送服務與網路訂餐。\n(7)本券商品不得與其他優惠合併使用且一桌限用一張。\n(8)本券商品無法累積會員積分。\n(9)本商品恕無法折抵商場停車。"
    beauty_default = "(1)使用本券請提前預約，預約時請告知使用本券及內容。\n(2)本券平假日皆適用。\n(3)本券商品無法累積會員積分。\n(4)本券不得指定美容師、療程房型；如您是敏感膚質，或有生理期、懷孕、心血管疾病等特殊身體狀況，請於課程開始前，主動告知美容師。\n(5)兌換數量依各門市現貨為準，若該門市已無存貨，請至鄰近門市兌換，或指定飲料可選擇更換其他系列品項。"
    entertainment_default = "(1)本優惠為單人使用，不建議 12 歲以下孩童參與。\n(2)本遊戲 4 人成團，未滿 8 人需與其他玩家併團遊戲。\n(3)本優惠活動時間約 100 分鐘(含講解)。"

    st.write("---")
    st.subheader("📦 商品品項與規格細節填寫")
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
                p_desc = st.text_area(f"商品描述及文案說明 ({i})", placeholder="關於商品的詳細說明", key=f"p_desc_{i}_{cid}", height=100)

            with p_c2:
                st.write("**📅 商品券兌換時間設定**")
                validity_type = st.radio(f"兌換類型 ({i})", ["常態商品 (系統自訂50天)", "季節商品 (自訂截止日)"], key=f"v_type_{i}_{cid}")
                t_now = datetime.now()
                if validity_type
