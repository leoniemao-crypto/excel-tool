


import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image
from datetime import datetime, timedelta
import io

st.title("📦 商品資料自動轉 Excel 工具")

# 1. 表單輸入欄位
prod_name = st.text_input("Product Name（商品名稱）")
uploaded_file = st.file_uploader("Main Image（主圖，會自動調整為 375x375）", type=["jpg", "jpeg", "png"])
instructions = st.text_area("Instructions for use（商品詳細說明，上限 32,767 字）")

if st.button("🚀 生成並下載 Excel"):
    if not prod_name:
        st.error("請至少輸入商品名稱！")
    else:
        # 2. 自動計算日期（今天 + 50 天）
        coupon_validity = (datetime.now() + timedelta(days=50)).strftime("%Y-%m-%d")
        
        # 3. 建立 Excel 檔案
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Product Data"
        
        # 寫入標題欄
        headers = ["Product Name", "Main Image", "Coupon Validity", "Instructions for use"]
        ws.append(headers)
        
        # 寫入文字資料
        ws["A2"] = prod_name
        ws["C2"] = coupon_validity
        ws["D2"] = instructions
        
        # 4. 處理圖片縮放 (375 x 375)
        if uploaded_file is not None:
            img = Image.open(uploaded_file)
            img = img.resize((375, 375))  # 強制設定規格
            
            # 將圖片轉為快取格式讓 Excel 讀取
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            
            xl_img = OpenpyxlImage(img_byte_arr)
            
            # 調整 Excel 儲存格寬高以符合 375x375 像素
            ws.row_dimensions[2].height = 285  # 375 像素約為 285 點
            ws.column_dimensions['B'].width = 50  # 欄寬調整
            
            ws.add_image(xl_img, 'B2')
        
        # 5. 匯出檔案
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.download_button(
            label="💾 點我下載 Excel 檔案",
            data=excel_buffer,
            file_name=f"{prod_name}_product_info.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Excel 產生成功！")