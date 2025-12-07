import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO

# --- 設定網頁標題 ---
st.set_page_config(page_title="自動化 PPT 生成器", page_icon="📊")
st.title("📊 自動化 PPT 生成器")
st.caption("輸入多組資料，一鍵生成包含多頁的 PowerPoint 簡報。")

# --- 初始化 Session State (用來暫存多頁資料) ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 側邊欄：輸入資料區域 ---
with st.sidebar:
    st.header("📝 新增頁面資料")
    st.info("請輸入第 4 步所需的欄位")
    
    # 輸入欄位
    case_info = st.text_input("1. 案號 / 日期", placeholder="例如：US 11,531,238 B2 / 2020.05.09")
    problem = st.text_area("2. 解決問題", placeholder="描述此專利解決了什麼技術問題...")
    spirit = st.text_area("3. 發明精神", placeholder="描述此發明的核心精神或技術手段...")
    key_point = st.text_input("4. 一句重點", placeholder="例如：第一與第二基板上的配向層方向相互垂直...")

    # 新增按鈕
    if st.button("➕ 加入此頁到簡報"):
        if case_info and problem and spirit and key_point:
            # 將資料存入 session_state
            st.session_state['slides_data'].append({
                "case_info": case_info,
                "problem": problem,
                "spirit": spirit,
                "key_point": key_point
            })
            st.success(f"已新增第 {len(st.session_state['slides_data'])} 頁！")
        else:
            st.warning("⚠️ 請將四個欄位都填寫完整。")

    # 清除所有資料按鈕
    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有頁面"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：顯示已輸入的資料與下載 ---

if not st.session_state['slides_data']:
    st.info("👈 請從左側側邊欄開始輸入資料，並點擊「加入此頁到簡報」。")
else:
    st.subheader(f"📋 預覽已輸入的 {len(st.session_state['slides_data'])} 頁內容")
    
    # 顯示目前已輸入的卡片
    for i, data in enumerate(st.session_state['slides_data']):
        with st.expander(f"第 {i+1} 頁：{data['case_info']}", expanded=False):
            st.markdown(f"**解決問題：** {data['problem']}")
            st.markdown(f"**發明精神：** {data['spirit']}")
            st.markdown(f"**一句重點：** {data['key_point']}")

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        # 設定為 16:9 寬螢幕 (預設是 4:3)
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            # 使用空白版型 (Layout 6 is usually blank)
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # --- 1. 案號 / 日期 (左上角紅框位置) ---
            # 位置估計: 左 2.5英吋, 上 1.2英吋
            left = Inches(2.5)
            top = Inches(1.2)
            width = Inches(4.0)
            height = Inches(0.8)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(14)
            p.font.bold = True

            # --- 2 & 3. 解決問題 與 發明精神 (中間區域) ---
            # 位置估計: 左 0.5英吋, 上 4.0英吋 (根據截圖大概位置)
            left = Inches(0.5)
            top = Inches(4.0)
            width = Inches(12.0)
            height = Inches(2.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            
            # 解決問題段落
            p1 = tf.add_paragraph()
            p1.text = "• 解決問題：" + data['problem']
            p1.font.size = Pt(16)
            p1.space_after = Pt(10) # 段落間距

            # 發明精神段落
            p2 = tf.add_paragraph()
            p2.text = "• 發明精神：" + data['spirit']
            p2.font.size = Pt(16)

            # --- 4. 一句重點 (底部長條) ---
            # 畫一個色塊當底圖
            left = Inches(0.5)
            top = Inches(6.5)
            width = Inches(12.3)
            height = Inches(0.8)
            
            # 新增矩形圖案
            from pptx.enum.shapes import MSO_SHAPE
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 192, 0) # 金黃色底 (類似截圖)
            shape.line.color.rgb = RGBColor(255, 192, 0) # 邊框同色

            # 在圖案中填字
            tf = shape.text_frame
            tf.vertical_anchor = 3 # MSO_ANCHOR.MIDDLE (垂直置中)
            p = tf.paragraphs[0]
            p.text = data['key_point']
            p.alignment = PP_ALIGN.CENTER # 水平置中
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0) # 黑色文字

        return prs

    # --- 下載按鈕 ---
    if st.button("🚀 生成 PowerPoint (.pptx)"):
        prs = generate_ppt(st.session_state['slides_data'])
        
        # 存到記憶體中
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        
        st.download_button(
            label="📥 點擊下載您的簡報",
            data=binary_output,
            file_name="generated_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
