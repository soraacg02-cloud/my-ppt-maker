import streamlit as st
import requests
from PIL import Image
from io import BytesIO

# --- 設定網頁標題與介面 ---
st.set_page_config(page_title="PPT 圖片生成器", page_icon="🖼️")
st.title("🖼️ PPT 圖片生成器")
st.caption("請在下方貼上圖片網址，系統將自動讀取。")

# --- 步驟 1: 圖片輸入 (貼上網址) ---
image_url = st.text_input("🌐 請貼上圖片網址 (Image URL)", placeholder="https://example.com/image.jpg")

# 建立一個變數來存放處理好的圖片
processed_image = None

if image_url:
    try:
        # 顯示讀取中的狀態
        with st.spinner("正在下載圖片..."):
            # 發送請求抓取圖片
            response = requests.get(image_url, timeout=10)
            response.raise_for_status() # 檢查網址是否有效 (404/500 等錯誤)
            
            # 將下載的資料轉為圖片格式
            processed_image = Image.open(BytesIO(response.content))
            
            # 顯示成功訊息與圖片
            st.success("圖片讀取成功！")
            st.image(processed_image, caption="預覽圖片", use_column_width=True)

    except requests.exceptions.MissingSchema:
        st.error("❌ 網址格式錯誤，請包含 http:// 或 https://")
    except requests.exceptions.ConnectionError:
        st.error("❌ 無法連線，請檢查網址是否正確。")
    except Exception as e:
        st.error(f"❌ 發生錯誤，無法讀取圖片：{e}")

# --- 步驟 2: 製作 PPT (範例功能) ---
st.divider() # 分隔線

if processed_image:
    st.subheader("🛠️ 製作選項")
    ppt_title = st.text_input("輸入 PPT 標題", "我的自動生成簡報")
    
    if st.button("🚀 開始製作 PPT"):
        # 這裡未來會放入製作 PPT 的程式碼
        # 目前先顯示成功動畫
        st.balloons()
        st.success(f"已針對「{ppt_title}」生成簡報！(這是示範功能)")
else:
    st.info("請先貼上有效的圖片網址，才能進行下一步。")
