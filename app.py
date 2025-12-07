import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from io import BytesIO

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (讀取舊 PPT 產出新格式)")
st.caption("上傳既有的 PPT 檔案，自動提取其中的圖片並重新排版。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：從 PPT 中提取圖片與文字 ---
def extract_data_from_pptx(uploaded_pptx):
    """
    讀取上傳的 PPT，回傳：
    1. 找到的第一張圖片的 binary data (若無則 None)
    2. 找到的所有文字內容 (字串)
    """
    try:
        prs = Presentation(uploaded_pptx)
        # 預設只讀取第一張投影片 (通常原始資料是一頁一案)
        slide = prs.slides[0]
        
        extracted_img = None
        extracted_text = []

        # 遍歷所有物件
        for shape in slide.shapes:
            # 1. 抓取圖片 (Shape Type 13 = PICTURE)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                # 只抓第一張找到的圖片 (假設最大那張就是主要的圖)
                if extracted_img is None:
                    extracted_img = shape.image.blob
            
            # 2. 抓取文字 (如果有 Text Frame)
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip():
                        extracted_text.append(paragraph.text.strip())
        
        return extracted_img, "\n".join(extracted_text)
    
    except Exception as e:
        st.error(f"解析 PPT 時發生錯誤: {e}")
        return None, ""

# --- 側邊欄：輸入資料區域 ---
with st.sidebar:
    st.header("1. 上傳原始資料")
    
    # 修改：上傳 PPTX 檔案
    uploaded_file = st.file_uploader(
        "上傳原始 PPT 檔案 (.pptx)", 
        type=['pptx'], 
        help="系統將會自動抓取此 PPT 內的第一張圖片作為圖示。"
    )

    # 暫存變數
    ppt_image_blob = None
    extracted_txt_content = ""

    if uploaded_file:
        with st.spinner("正在分析 PPT 內容..."):
            ppt_image_blob, extracted_txt_content = extract_data_from_pptx(uploaded_file)
            
            if ppt_image_blob:
                st.success("✅ 已成功提取圖片！")
                st.image(ppt_image_blob, caption="從 PPT 提取的圖片", use_column_width=True)
            else:
                st.warning("⚠️ 此 PPT 中找不到圖片。")

            # 顯示提取的文字供參考
            with st.expander("🔍 查看 PPT 內的文字 (可複製)", expanded=True):
                st.text_area("原始文字內容", extracted_txt_content, height=150)

    st.divider()
    st.header("2. 填寫排版內容")
    st.info("請參考上方提取的文字，填入下欄：")

    # 輸入欄位
    case_info = st.text_input("案號 / 日期", placeholder="例如：US 11,531,238 B2 / 2020.05.09")
    problem = st.text_area("解決問題", placeholder="描述此專利解決了什麼技術問題...")
    spirit = st.text_area("發明精神", placeholder="描述此發明的核心精神或技術手段...")
    key_point = st.text_input("一句重點", placeholder="例如：第一與第二基板上的配向層方向相互垂直...")

    # 新增按鈕
    if st.button("➕ 加入此頁到簡報", type="primary"):
        if case_info and problem and spirit and key_point:
            
            # 圖片處理：優先使用從 PPT 抓到的圖
            image_data_to_save = ppt_image_blob
            image_name_str = uploaded_file.name if uploaded_file else "無圖片"

            # 將資料存入 session_state
            st.session_state['slides_data'].append({
                "case_info": case_info,
                "problem": problem,
                "spirit": spirit,
                "key_point": key_point,
                "image_data": image_data_to_save,
                "image_name": image_name_str
            })
            st.success(f"已新增第 {len(st.session_state['slides_data'])} 頁！")
        else:
            st.warning("⚠️ 請將四個文字欄位都填寫完整。")

    # 清除所有資料按鈕
    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有頁面"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：顯示已輸入的資料與下載 ---

if not st.session_state['slides_data']:
    st.info("👈 請從左側開始：先上傳 PPT，系統會自動抓圖，您只需填寫文字。")
else:
    st.subheader(f"📋 預覽已輸入的 {len(st.session_state['slides_data'])} 頁內容")
    
    col_count = 0
    cols = st.columns(3)
    
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[col_count % 3]:
            with st.container(border=True):
                st.markdown(f"#### 第 {i+1} 頁")
                st.text(f"來源：{data['image_name']}")
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    st.markdown("*[未偵測到圖片]*")
                st.markdown(f"**重點：** {data['key_point']}")
        col_count += 1

    st.divider()

    # --- PPT 生成邏輯 (保持不變，負責排版) ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白版型

            # 1. 案號 (左上)
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(1.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(24)
            p.font.bold = True
            
            # 2. 圖片 (右上 - 使用從 PPT 提取的資料)
            if data['image_data']:
                image_stream = BytesIO(data['image_data'])
                # 設定位置與高度限制
                slide.shapes.add_picture(image_stream, Inches(5.5), Inches(0.5), height=Inches(4.0))

            # 3. 文字區 (中下)
            left, top, width, height = Inches(0.5), Inches(4.8), Inches(12.3), Inches(1.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            
            p1 = tf.add_paragraph()
            p1.text = "• 解決問題：" + data['problem']
            p1.font.size = Pt(18)
            p1.space_after = Pt(12)

            p2 = tf.add_paragraph()
            p2.text = "• 發明精神：" + data['spirit']
            p2.font.size = Pt(18)

            # 4. 重點 (底部黃底)
            left, top, width, height = Inches(0.5), Inches(6.5), Inches(12.3), Inches(0.8)
            shape = slide.shapes.add_shape(MSO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(255, 192, 0)
            shape.line.color.rgb = RGBColor(255, 192, 0)

            p = shape.text_frame.paragraphs[0]
            p.text = data['key_point']
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 0, 0)
            shape.text_frame.vertical_anchor = 3

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
            file_name="organized_patent_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
