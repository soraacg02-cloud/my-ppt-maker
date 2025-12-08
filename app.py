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
st.set_page_config(page_title="PPT 重組生成器 (智慧搜圖版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (智慧圖號提取版)")
st.caption("修正：自動從詳細的代表圖說明中提取「圖號」(如 FIG. 3E)，解決因說明文字過長導致搜圖失敗的問題。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：依據關鍵字搜尋 PDF 並截圖 (升級版) ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    """
    從 target_fig_text (代表圖說明) 中提取圖號，並在 PDF 中搜尋該圖號。
    """
    if not target_fig_text:
        return None, "無文字"

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        # --- 步驟 1: 智慧提取圖號 ---
        # 我們不要拿整行去搜，只抓像 "FIG. 3E", "Figure 1", "圖2" 這樣的關鍵字
        # Regex 解釋:
        # (?:FIG\.?|Figure|圖)  -> 匹配 FIG. 或 FIG 或 Figure 或 圖 (不分大小寫)
        # \s* -> 允許中間有空白
        # [0-9]+                -> 數字
        # [A-Za-z]* -> 可選的英文後綴 (如 3E 的 E)
        pattern = re.compile(r'((?:FIG\.?|Figure|圖)\s*[0-9]+[A-Za-z]*)', re.IGNORECASE)
        
        search_keywords = []
        lines = target_fig_text.split('\n')
        
        # 掃描每一行，找出所有可能的圖號
        for line in lines:
            match = pattern.search(line)
            if match:
                # 抓到了！例如 "FIG. 3E"
                # 去除空白，標準化 (例如 "FIG. 3E" -> "FIG.3E") 以利比對
                raw_keyword = match.group(1)
                clean_keyword = raw_keyword.replace(" ", "").upper()
                search_keywords.append(clean_keyword)
        
        # 如果 Regex 沒抓到 (例如使用者只寫 "參考下圖")，只好用第一行的前10個字試試看
        if not search_keywords:
             first_line = lines[0].strip()
             if first_line:
                 search_keywords.append(first_line[:10].replace(" ", "").upper())

        # --- 步驟 2: 在 PDF 中搜尋 ---
        found_page_index = None
        matched_keyword_log = ""

        # 優先搜尋提取到的第一個圖號 (通常代表圖是第一個提到的)
        target_keyword = search_keywords[0] if search_keywords else ""
        
        if not target_keyword:
            return None, "無法識別圖號"

        for i, page in enumerate(doc):
            page_text = page.get_text()
            # 移除空白與轉大寫來比對
            clean_page_text = page_text.replace(" ", "").upper()
            
            if target_keyword in clean_page_text:
                found_page_index = i
                matched_keyword_log = target_keyword
                break
        
        if found_page_index is not None:
            page = doc[found_page_index]
            mat = fitz.Matrix(2, 2) # 放大 2 倍
            pix = page.get_pixmap(matrix=mat)
            return pix.tobytes("png"), f"成功匹配: {matched_keyword_log}"
            
        return None, f"PDF中找不到: {target_keyword}"

    except Exception as e:
        print(f"PDF 解析錯誤: {e}")
        return None, f"錯誤: {str(e)}"

# --- 函數：提取專利號 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    # 支援 CN, TW, TWI, US 等格式
    match = re.search(r'([a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match:
        return match.group(1)
    return ""

# --- 函數：解析 Word 檔案 (狀態機邏輯) ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        current_case = {
            "case_info": "", 
            "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
            "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
        }
        current_field = None 
        
        debug_raw_lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue
            
            # --- 1. 新案件判斷 (最高優先) ---
            if "案號" in text or "索號" in text:
                if current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                
                current_field = "case_info_block"
                current_case["case_info"] = text 
                extracted_no = extract_patent_number_from_text(text)
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no
                continue

            # --- 2. 欄位切換 ---
            if "解決問題" in text:
                current_field = "problem"
                content = re.sub(r'^[0-9.．]*\s*解決問題[:：]?\s*', '', text)
                current_case["problem"] = content
                continue

            elif "發明精神" in text:
                current_field = "spirit"
                content = re.sub(r'^[0-9.．]*\s*發明精神[:：]?\s*', '', text)
                current_case["spirit"] = content
                continue

            elif "重點" in text:
                current_field = "key_point"
                content = re.sub(r'^[0-9.．]*\s*(一句)?重點[:：]?\s*', '', text)
                current_case["key_point"] = content
                continue

            elif "代表圖" in text:
                current_field = "rep_fig"
                content = re.sub(r'^[0-9.．]*\s*代表圖[:：]?\s*', '', text).strip()
                current_case["rep_fig_text"] = content
                continue

            # --- 3. 內容填充 ---
            if current_field == "case_info_block":
                current_case["case_info"] += "\n" + text
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no

            elif current_field == "rep_fig":
                current_case["rep_fig_text"] += "\n" + text

            elif current_field == "problem":
                current_case["problem"] += "\n" + text

            elif current_field == "spirit":
                current_case["spirit"] += "\n" + text

            elif current_field == "key_point":
                current_case["key_point"] += "\n" + text

        if current_case["case_info"]:
            cases.append(current_case)
            
        return cases, debug_raw_lines

    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return [], []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    word_file = st.file_uploader("Word 檔案 (.docx)", type=['docx'])
    pdf_files = st.file_uploader("PDF 檔案 (.pdf)", type=['pdf'], accept_multiple_files=True)
    
    if word_file and st.button("🔄 開始智能整合", type="primary"):
        extracted_cases, raw_lines = parse_word_file(word_file)
        
        # 讀取 PDF
        pdf_file_map = {}
        if pdf_files:
            for pdf in pdf_files:
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                pdf_file_map[clean_name] = pdf.read()

        match_count = 0
        
        with st.spinner("正在搜尋圖片..."):
            for case in extracted_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"]
                
                matched_pdf_bytes = None
                matched_pdf_name = ""
                
                # 尋找對應 PDF
                for pdf_key, pdf_bytes in pdf_file_map.items():
                    if case_key and ((pdf_key.lower() in case_key.lower()) or (case_key.lower() in pdf_key.lower())):
                        if len(case_key) > 4: 
                            matched_pdf_bytes = pdf_bytes
                            matched_pdf_name = pdf_key
                            break
                
                if matched_pdf_bytes and target_fig:
                    img_data, log_msg = extract_specific_figure_from_pdf(matched_pdf_bytes, target_fig)
                    if img_data:
                        case["image_data"] = img_data
                        case["image_name"] = f"截取成功 ({matched_pdf_name})"
                        match_count += 1
                    else:
                        case["image_name"] = f"找不到圖 ({log_msg})"
                else:
                    if not matched_pdf_bytes:
                        case["image_name"] = "無對應 PDF"
                    else:
                        case["image_name"] = "Word 無代表圖資訊"

        if extracted_cases:
            st.session_state['slides_data'].extend(extracted_cases)
            st.success(f"處理完成！共 {len(extracted_cases)} 筆，截取 {match_count} 張圖。")
        else:
            st.warning("Word 解析無資料。")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面 ---
if not st.session_state['slides_data']:
    st.info("👈 請上傳檔案。此版本能自動從長篇說明中抓取「FIG. 3E」作為搜圖關鍵字。")
else:
    st.subheader(f"📋 預覽")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.text(data['case_info'])
                
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    # 顯示文字內容
                    display_text = data['rep_fig_text'] if data['rep_fig_text'].strip() else "(Word中無代表圖資訊)"
                    st.warning(f"無圖片 ({data['image_name']})，將填入文字：\n{display_text[:50]}...")
                
                st.caption(f"重點：{data['key_point']}")

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 1. 左上：案號 / 日期 / 公司 (條列式)
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(2.0)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True

            info_lines = data['case_info'].split('\n')
            for line in info_lines:
                if line.strip():
                    p = tf.add_paragraph()
                    p.text = line.strip()
                    p.font.size = Pt(20)
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(0, 0, 0)
                    p.alignment = PP_ALIGN.LEFT
            
            # 2. 右上：綠框區域
            img_left = Inches(5.5)
            img_top = Inches(0.5)
            img_height = Inches(4.0)
            img_width = Inches(7.0)

            if data['image_data']:
                image_stream = BytesIO(data['image_data'])
                slide.shapes.add_picture(image_stream, img_left, img_top, height=img_height)
            else:
                # 替代文字 (16pt 條列式)
                txBox = slide.shapes.add_textbox(img_left, img_top, img_width, img_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                
                content_text = data['rep_fig_text'] if data['rep_fig_text'].strip() else "(Word中無代表圖資訊)"
                
                lines = content_text.split('\n')
                for line in lines:
                    if line.strip():
                        p = tf.add_paragraph()
                        p.text = line.strip()
                        p.font.size = Pt(16)
                        p.font.bold = False
                        p.alignment = PP_ALIGN.LEFT

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
            shape.text_frame.vertical_anchor = MSO_SHAPE.RECTANGLE

        return prs

    if st.button("🚀 生成 PowerPoint (.pptx)", type="primary"):
        prs = generate_ppt(st.session_state['slides_data'])
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        
        st.download_button(
            label="📥 下載 PPT",
            data=binary_output,
            file_name="smart_figure_search_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
