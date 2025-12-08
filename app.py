import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from io import BytesIO
import docx
import fitz  # PyMuPDF
import re

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器")
st.caption("優化：綠框區文字改為 16pt 條列式顯示，完整保留 Word 換行格式。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：依據關鍵字搜尋 PDF 並截圖 ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    """
    在 PDF 中搜尋 target_fig_text。
    """
    if not target_fig_text:
        return None

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        # 只取第一行作為搜尋關鍵字 (避免 Word 裡寫了多行描述導致搜尋失敗)
        search_keyword = target_fig_text.split('\n')[0].strip()
        clean_target = search_keyword.replace(" ", "")
        
        found_page_index = None

        for i, page in enumerate(doc):
            page_text = page.get_text()
            clean_page_text = page_text.replace(" ", "")
            if clean_target in clean_page_text:
                found_page_index = i
                break
        
        if found_page_index is not None:
            page = doc[found_page_index]
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)
            return pix.tobytes("png")
            
        return None

    except Exception as e:
        print(f"PDF 解析錯誤: {e}")
        return None

# --- 函數：提取專利號 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match:
        return match.group(1)
    return ""

# --- 函數：解析 Word 檔案 (修正換行處理) ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
            "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
        }
        current_field = None 

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            # --- 關鍵字判斷 ---
            if any(k in text for k in ["案號", "日期", "申請日", "索號"]):
                if ("案號" in text or "索號" in text) and current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                current_field = "case_info_block"
                current_case["case_info"] += text + "\n"
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no

            elif "解決問題" in text:
                current_field = "problem"
                current_case["problem"] = text.replace("解決問題", "").replace(":", "").replace("：", "").strip()

            elif "發明精神" in text:
                current_field = "spirit"
                current_case["spirit"] = text.replace("發明精神", "").replace(":", "").replace("：", "").strip()

            elif "重點" in text:
                current_field = "key_point"
                current_case["key_point"] = text.replace("一句重點", "").replace("重點", "").replace(":", "").replace("：", "").strip()

            elif "代表圖" in text:
                current_field = "rep_fig"
                # 清理標題，只留內容
                clean_fig = text.replace("5", "").replace(".", "").replace("代表圖", "").replace(":", "").replace("：", "").strip()
                current_case["rep_fig_text"] = clean_fig

            else:
                # 續行文字處理 (重點：加入 \n 保留換行)
                if current_field == "case_info_block":
                    current_case["case_info"] += text + "\n"
                    extracted_no = extract_patent_number_from_text(current_case["case_info"])
                    if extracted_no:
                        current_case["raw_case_no"] = extracted_no
                elif current_field in ["problem", "spirit", "key_point"]:
                    current_case[current_field] += "\n" + text
                elif current_field == "rep_fig":
                    # 這裡加入換行符號，確保 Word 裡的條列式能被保留
                    current_case["rep_fig_text"] += "\n" + text 

        if current_case["case_info"]:
            cases.append(current_case)
        return cases
    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    word_file = st.file_uploader("Word 檔案 (.docx)", type=['docx'])
    pdf_files = st.file_uploader("PDF 檔案 (.pdf)", type=['pdf'], accept_multiple_files=True)
    
    if word_file and st.button("🔄 開始智能整合", type="primary"):
        # 1. 解析 Word
        extracted_cases = parse_word_file(word_file)
        
        # 2. 讀取 PDF
        pdf_file_map = {}
        if pdf_files:
            for pdf in pdf_files:
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                pdf_file_map[clean_name] = pdf.read()

        # 3. 配對
        match_count = 0
        debug_logs = []
        
        with st.spinner("正在處理..."):
            for case in extracted_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"]
                
                matched_pdf_bytes = None
                matched_name = ""
                
                for pdf_key, pdf_bytes in pdf_file_map.items():
                    if case_key and ((pdf_key.lower() in case_key.lower()) or (case_key.lower() in pdf_key.lower())):
                        if len(case_key) > 4: 
                            matched_pdf_bytes = pdf_bytes
                            matched_name = pdf_key
                            break
                
                if matched_pdf_bytes and target_fig:
                    img_data = extract_specific_figure_from_pdf(matched_pdf_bytes, target_fig)
                    if img_data:
                        case["image_data"] = img_data
                        case["image_name"] = f"成功截取: {target_fig}"
                        match_count += 1
                    else:
                        case["image_name"] = f"找不到圖，使用文字"
                else:
                    case["image_name"] = "無對應資料"

        if extracted_cases:
            st.session_state['slides_data'].extend(extracted_cases)
            st.success(f"處理完成！共 {len(extracted_cases)} 筆。")
        else:
            st.warning("Word 解析無資料。")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面 ---
if not st.session_state['slides_data']:
    st.info("👈 請上傳檔案。")
else:
    st.subheader(f"📋 預覽")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.caption(data['raw_case_no'])
                
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    # 預覽顯示文字內容
                    st.info(f"無圖片，將填入：\n\n{data['rep_fig_text']}")
                
                st.caption(f"重點：{data['key_point']}")

    st.divider()

    # --- PPT 生成邏輯 (修正字體與排版) ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 1. 左上：案號
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(1.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(20)
            p.font.bold = True
            
            # 2. 右上：綠框區域 (圖片 或 文字)
            img_left = Inches(5.5)
            img_top = Inches(0.5)
            img_height = Inches(4.0)
            img_width = Inches(7.0)

            if data['image_data']:
                # 有圖 -> 放圖
                image_stream = BytesIO(data['image_data'])
                slide.shapes.add_picture(image_stream, img_left, img_top, height=img_height)
            else:
                # 沒圖 -> 放文字 (修正點：字體 16pt，條列式)
                txBox = slide.shapes.add_textbox(img_left, img_top, img_width, img_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                
                # 將文字內容依據換行符號切割，逐行加入
                lines = (data['rep_fig_text'] if data['rep_fig_text'] else "(未指定)").split('\n')
                
                for line in lines:
                    if line.strip(): # 避免空白行
                        p = tf.add_paragraph()
                        p.text = line.strip()
                        p.font.size = Pt(16) # 設定為 16pt
                        p.font.bold = False
                        p.font.color.rgb = RGBColor(0, 0, 0) # 黑色文字
                        p.alignment = PP_ALIGN.LEFT # 靠左對齊 (條列式風格)
                
                # 設定垂直對齊 (預設是靠上，如果要垂直置中可解開下面這行)
                # txBox.text_frame.vertical_anchor = MSO_SHAPE.RECTANGLE

            # 3. 中下：文字區
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

            # 4. 底部：重點
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
            file_name="final_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
