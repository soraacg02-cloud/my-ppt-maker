import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from io import BytesIO
import docx # 引入 python-docx 用來讀取 Word 檔

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器", page_icon="📑", layout="wide")
st.title("📑 PPT 自動化生成器 (支援 Word/PPT 匯入)")
st.caption("支援多來源匯入：可上傳 Word 自動拆解多案，或上傳 PPT 提取圖文。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數 1：從 PPT 中提取圖片與文字 (既有功能) ---
def extract_data_from_pptx(uploaded_pptx):
    try:
        prs = Presentation(uploaded_pptx)
        slide = prs.slides[0] # 預設只讀第一頁
        extracted_img = None
        extracted_text = []

        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if extracted_img is None:
                    extracted_img = shape.image.blob
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.text.strip():
                        extracted_text.append(paragraph.text.strip())
        return extracted_img, "\n".join(extracted_text)
    except Exception as e:
        st.error(f"解析 PPT 時發生錯誤: {e}")
        return None, ""

# --- 函數 2：從 Word 中批次提取多案資料 (新增功能) ---
def parse_word_file(uploaded_docx):
    """
    解析 Word 檔案，依據關鍵字自動拆解成多筆資料。
    假設格式為：
    案號：xxx
    解決問題：yyy
    發明精神：zzz
    一句重點：aaa
    (重複循環)
    """
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        # 初始化一個暫存的案子資料
        current_case = {"case_info": "", "problem": "", "spirit": "", "key_point": "", "image_data": None, "image_name": "Word匯入"}
        current_field = None # 記錄目前正在讀取哪個欄位

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue # 跳過空行

            # 判斷關鍵字 (支援常見寫法)
            if "案號" in text or "日期" in text:
                # 如果已經有資料且又讀到「案號」，代表是下一筆案子，先儲存上一筆
                if current_case["case_info"] and current_field != "case_info":
                    cases.append(current_case)
                    current_case = {"case_info": "", "problem": "", "spirit": "", "key_point": "", "image_data": None, "image_name": "Word匯入"}
                
                current_field = "case_info"
                # 去除標籤文字
                clean_text = text.replace("案號", "").replace("日期", "").replace("/", "").replace(":", "").replace("：", "").strip()
                current_case["case_info"] = clean_text

            elif "解決問題" in text:
                current_field = "problem"
                clean_text = text.replace("解決問題", "").replace(":", "").replace("：", "").strip()
                current_case["problem"] = clean_text

            elif "發明精神" in text:
                current_field = "spirit"
                clean_text = text.replace("發明精神", "").replace(":", "").replace("：", "").strip()
                current_case["spirit"] = clean_text

            elif "重點" in text:
                current_field = "key_point"
                clean_text = text.replace("一句重點", "").replace("重點", "").replace(":", "").replace("：", "").strip()
                current_case["key_point"] = clean_text

            else:
                # 如果該行沒有關鍵字，但目前正在某個欄位中，則視為該欄位的續行 (多行文字)
                if current_field:
                    current_case[current_field] += "\n" + text

        # 迴圈結束後，別忘了存最後一筆
        if current_case["case_info"]:
            cases.append(current_case)
        
        return cases

    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料來源")
    
    # 頁籤：選擇匯入方式
    import_mode = st.radio("選擇匯入方式", ["手動輸入 / PPT 提取", "Word 批次匯入"])

    if import_mode == "Word 批次匯入":
        st.info("請上傳 Word 檔 (.docx)，系統將依據「案號」、「解決問題」等關鍵字自動分頁。")
        word_file = st.file_uploader("上傳 Word 檔案", type=['docx'])
        
        if word_file:
            if st.button("🔄 開始解析 Word", type="primary"):
                extracted_cases = parse_word_file(word_file)
                if extracted_cases:
                    st.session_state['slides_data'].extend(extracted_cases)
                    st.success(f"成功匯入 {len(extracted_cases)} 筆資料！請看右側預覽。")
                else:
                    st.warning("未找到有效資料，請確認 Word 內容包含「案號」、「解決問題」等關鍵字。")

    else:
        # --- 原有的 PPT / 手動輸入模式 ---
        uploaded_file = st.file_uploader(
            "上傳原始 PPT (.pptx) 以提取圖文", 
            type=['pptx'],
            help="自動抓取 PPT 第一張圖片與文字。"
        )

        ppt_image_blob = None
        extracted_txt_content = ""

        if uploaded_file:
            with st.spinner("分析 PPT 中..."):
                ppt_image_blob, extracted_txt_content = extract_data_from_pptx(uploaded_file)
                if ppt_image_blob:
                    st.success("已提取圖片")
                    st.image(ppt_image_blob, caption="PPT 圖片", use_column_width=True)
                
                with st.expander("查看 PPT 文字", expanded=True):
                    st.text_area("內容", extracted_txt_content, height=100)

        st.divider()
        st.header("2. 編輯內容")
        case_info = st.text_input("案號 / 日期")
        problem = st.text_area("解決問題")
        spirit = st.text_area("發明精神")
        key_point = st.text_input("一句重點")

        if st.button("➕ 加入此頁到簡報", type="primary"):
            if case_info and problem and spirit and key_point:
                image_data = ppt_image_blob
                image_name = uploaded_file.name if uploaded_file else "無圖片"
                
                st.session_state['slides_data'].append({
                    "case_info": case_info,
                    "problem": problem,
                    "spirit": spirit,
                    "key_point": key_point,
                    "image_data": image_data,
                    "image_name": image_name
                })
                st.success("已新增頁面！")
            else:
                st.warning("請填寫所有欄位。")

    # 清除按鈕
    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有頁面"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：預覽與下載 ---

if not st.session_state['slides_data']:
    st.info("👈 請從左側開始匯入資料。")
else:
    st.subheader(f"📋 預覽 ({len(st.session_state['slides_data'])} 頁)")
    
    col_count = 0
    cols = st.columns(3)
    
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[col_count % 3]:
            with st.container(border=True):
                st.markdown(f"#### 第 {i+1} 頁")
                st.text(f"案號：{data['case_info']}")
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    st.markdown("*(無圖片)*")
                st.markdown(f"**重點：** {data['key_point']}")
        col_count += 1

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 1. 案號 (左上)
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(1.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(24)
            p.font.bold = True
            
            # 2. 圖片 (右上)
            if data['image_data']:
                image_stream = BytesIO(data['image_data'])
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
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
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

    if st.button("🚀 生成 PowerPoint (.pptx)", type="primary"):
        prs = generate_ppt(st.session_state['slides_data'])
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        
        st.download_button(
            label="📥 下載 PPT",
            data=binary_output,
            file_name="auto_generated_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
