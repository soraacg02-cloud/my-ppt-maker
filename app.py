import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO

# --- 設定網頁標題 ---
st.set_page_config(page_title="自動化 PPT 生成器", page_icon="📊", layout="wide")
st.title("📊 自動化 PPT 生成器 (含圖片上傳)")
st.caption("輸入文字並上傳圖片，一鍵生成包含多頁的 PowerPoint 簡報。")

# --- 初始化 Session State (用來暫存多頁資料) ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 側邊欄：輸入資料區域 ---
with st.sidebar:
    st.header("📝 新增頁面資料")
    
    # 輸入欄位
    case_info = st.text_input("1. 案號 / 日期", placeholder="例如：US 11,531,238 B2 / 2020.05.09")
    
    # 新增：圖片上傳欄位
    uploaded_file = st.file_uploader("2. 上傳圖案 (綠框位置)", type=['png', 'jpg', 'jpeg'], help="請上傳圖片檔案，這會被放在 PPT 的右上位置")
    
    problem = st.text_area("3. 解決問題", placeholder="描述此專利解決了什麼技術問題...")
    spirit = st.text_area("4. 發明精神", placeholder="描述此發明的核心精神或技術手段...")
    key_point = st.text_input("5. 一句重點", placeholder="例如：第一與第二基板上的配向層方向相互垂直...")

    # 新增按鈕
    if st.button("➕ 加入此頁到簡報", type="primary"):
        if case_info and problem and spirit and key_point:
            
            # 處理圖片：如果有上傳，轉為二進位資料儲存
            image_data = None
            if uploaded_file is not None:
                image_data = uploaded_file.getvalue()
                uploaded_filename = uploaded_file.name
            else:
                uploaded_filename = "無圖片"

            # 將資料存入 session_state
            st.session_state['slides_data'].append({
                "case_info": case_info,
                "problem": problem,
                "spirit": spirit,
                "key_point": key_point,
                "image_data": image_data, # 儲存圖片資料
                "image_name": uploaded_filename
            })
            st.success(f"已新增第 {len(st.session_state['slides_data'])} 頁！")
        else:
            st.warning("⚠️ 請將所有文字欄位填寫完整 (圖片為選填)。")

    # 清除所有資料按鈕
    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有頁面"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：顯示已輸入的資料與下載 ---

if not st.session_state['slides_data']:
    st.info("👈 請從左側側邊欄開始輸入資料。若您的資料在 PPT 裡，請先將該圖示「另存成圖片」或「截圖」後上傳。")
else:
    st.subheader(f"📋 預覽已輸入的 {len(st.session_state['slides_data'])} 頁內容")
    
    # 顯示目前已輸入的卡片
    col_count = 0
    cols = st.columns(3) # 用三欄排列預覽
    
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[col_count % 3]:
            with st.container(border=True):
                st.markdown(f"#### 第 {i+1} 頁")
                st.text(f"案號：{data['case_info']}")
                if data['image_data']:
                    st.image(data['image_data'], caption=f"圖案：{data['image_name']}", use_column_width=True)
                else:
                    st.markdown("*[無圖片]*")
                st.markdown(f"**重點：** {data['key_point']}")
        col_count += 1

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        # 設定為 16:9 寬螢幕
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            # 使用空白版型 (Layout 6)
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # --- 1. 案號 / 日期 (左上角) ---
            # 位置: 左 0.5, 上 0.5
            left = Inches(0.5)
            top = Inches(0.5)
            width = Inches(5.0)
            height = Inches(1.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(24) # 加大字體
            p.font.bold = True
            
            # --- 2. 圖片 (綠框位置 - 右上/中) ---
            if data['image_data']:
                # 將二進位資料轉回串流以供 pptx 讀取
                image_stream = BytesIO(data['image_data'])
                
                # 位置設定 (根據截圖綠框位置)
                # 放在右半邊，留一點邊界
                img_left = Inches(5.5) 
                img_top = Inches(0.5)
                img_width = Inches(7.0) # 寬度設大一點
                img_height = Inches(4.0) # 高度限制
                
                # add_picture 可以只指定寬度或高度，另一個會自動等比例縮放
                # 這裡我們先限制高度，避免蓋到下面的文字
                slide.shapes.add_picture(image_stream, img_left, img_top, height=img_height)

            # --- 3. 解決問題 與 發明精神 (下方文字區 - 紅框) ---
            # 位置: 在圖片下方，約 5.0 英吋位置開始
            left = Inches(0.5)
            top = Inches(4.8) 
            width = Inches(12.3)
            height = Inches(1.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            
            # 解決問題
            p1 = tf.add_paragraph()
            p1.text = "• 解決問題：" + data['problem']
            p1.font.size = Pt(18)
            p1.space_after = Pt(12)

            # 發明精神
            p2 = tf.add_paragraph()
            p2.text = "• 發明精神：" + data['spirit']
            p2.font.size = Pt(18)

            # --- 4. 一句重點 (底部長條 - 黃底) ---
            left = Inches(0.5)
            top = Inches(6.5) # 底部
            width = Inches(12.3)
            height = Inches(0.8)
            
            from pptx.enum.shapes import MSO_SHAPE
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 192, 0) # 金黃色
            shape.line.color.rgb = RGBColor(255, 192, 0)

            tf = shape.text_frame
            tf.vertical_anchor = 3 # 垂直置中
            p = tf.paragraphs[0]
            p.text = data['key_point']
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)

        return prs

    # --- 下載按鈕 ---
    if st.button("🚀 生成 PowerPoint (.pptx)", type="primary"):
        prs = generate_ppt(st.session_state['slides_data'])
        
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        
        st.download_button(
            label="📥 點擊下載您的簡報",
            data=binary_output,
            file_name="patent_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
