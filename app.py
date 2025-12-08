import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from io import BytesIO
import docx
from docx.document import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import fitz  # PyMuPDF
import re

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器 (除錯最終版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (邏輯重構版)")
st.caption("重構：採用平鋪式迴圈處理，並新增「單案歷程」以追蹤文字歸類狀況。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 輔助函數：遍歷 Word 所有區塊 (含表格) ---
def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    else:
        raise ValueError("只支援讀取整份 Document")
    for child in parent_elm.iterchildren():
        if child.tag.endswith('p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('tbl'):
            yield Table(child, parent)

# --- 函數：依據關鍵字搜尋 PDF 並截圖 ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    if not target_fig_text:
        return None, "無文字"

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        # 智慧提取圖號 Regex
        pattern = re.compile(r'((?:FIG\.?|Figure|圖)\s*[0-9]+[A-Za-z]*)', re.IGNORECASE)
        
        search_keywords = []
        lines = target_fig_text.split('\n')
        for line in lines:
            match = pattern.search(line)
            if match:
                raw_keyword = match.group(1)
                clean_keyword = raw_keyword.replace(" ", "").upper()
                search_keywords.append(clean_keyword)
        
        if not search_keywords:
             first_line = lines[0].strip()
             if first_line:
                 search_keywords.append(first_line[:10].replace(" ", "").upper())

        target_keyword = search_keywords[0] if search_keywords else ""
        if not target_keyword:
            return None, "無法識別圖號"

        found_page_index = None
        matched_keyword_log = ""

        for i, page in enumerate(doc):
            page_text = page.get_text()
            clean_page_text = page_text.replace(" ", "").upper()
            if target_keyword in clean_page_text:
                found_page_index = i
                matched_keyword_log = target_keyword
                break
        
        if found_page_index is not None:
            page = doc[found_page_index]
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)
            return pix.tobytes("png"), f"成功匹配: {matched_keyword_log}"
            
        return None, f"PDF中找不到: {target_keyword}"
    except Exception as e:
        return None, f"錯誤: {str(e)}"

# --- 函數：提取專利號 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match:
        return match.group(1)
    return ""

# --- 函數：解析 Word 檔案 (平鋪邏輯 + 歷程記錄) ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        
        # 初始化當前案件
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
            "image_data": None, "image_name": "Word匯入", "raw_case_no": "",
            "parse_log": [] # 新增：記錄這筆案子吃到了哪些行
        }
        current_field = None 
        debug_raw_lines = [] # 全域除錯

        # --- 開始遍歷 ---
        for block in iter_block_items(doc):
            text = ""
            if isinstance(block, Paragraph):
                text = block.text.strip()
            elif isinstance(block, Table):
                # 簡單將表格內容轉為多行文字
                cell_texts = []
                for row in block.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text.strip():
                                cell_texts.append(p.text.strip())
                # 這裡我們將表格內容視為一個大的文字區塊處理，或者您可以選擇逐行處理
                # 為了邏輯簡單，我們把表格拆解成虛擬的行
                # 但這裡為了配合迴圈結構，我們需要一個小技巧：直接處理這些文字
                for cell_text in cell_texts:
                    # 遞迴呼叫邏輯太複雜，這裡直接複製貼上核心邏輯 (或封裝成不含狀態的函數)
                    # 為求保險，我們把表格文字插入到 text 處理流程中
                    # 但因為 python 迴圈特性，我們改為收集所有文字行再統一跑迴圈
                    pass 
                # 修正：為了支援表格，我們改為先收集所有 lines，再跑狀態機
            
        # --- 步驟 1: 將文檔完全平展為 Lines (解決表格/段落混合問題) ---
        all_lines = []
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                if block.text.strip():
                    all_lines.append(block.text.strip())
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if p.text.strip():
                                all_lines.append(p.text.strip())
        
        # --- 步驟 2: 狀態機迴圈 ---
        for text in all_lines:
            debug_raw_lines.append(text[:20]) # 記錄前20字

            # A. 判斷新案件 (案號/索號)
            if "案號" in text or "索號" in text:
                # 存檔上一筆
                if current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    # 開新的一筆
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": "",
                        "parse_log": []
                    }
                
                current_field = "case_info_block"
                current_case["case_info"] = text
                current_case["parse_log"].append(f"[Info] {text}")
                
                extracted_no = extract_patent_number_from_text(text)
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no
                continue

            # B. 判斷欄位切換
            if "解決問題" in text:
                current_field = "problem"
                content = re.sub(r'^[0-9.．]*\s*解決問題[:：]?\s*', '', text)
                current_case["problem"] = content
                current_case["parse_log"].append(f"[Problem Header] {text}")
                continue

            elif "發明精神" in text:
                current_field = "spirit"
                content = re.sub(r'^[0-9.．]*\s*發明精神[:：]?\s*', '', text)
                current_case["spirit"] = content
                current_case["parse_log"].append(f"[Spirit Header] {text}")
                continue

            elif "重點" in text:
                current_field = "key_point"
                content = re.sub(r'^[0-9.．]*\s*(一句)?重點[:：]?\s*', '', text)
                current_case["key_point"] = content
                current_case["parse_log"].append(f"[KeyPoint Header] {text}")
                continue

            elif "代表圖" in text:
                current_field = "rep_fig"
                content = re.sub(r'^[0-9.．]*\s*代表圖[:：]?\s*', '', text).strip()
                current_case["rep_fig_text"] = content
                current_case["parse_log"].append(f"[RepFig Header] {text}")
                continue

            # C. 內容填充 (狀態延續)
            if current_field == "case_info_block":
                current_case["case_info"] += "\n" + text
                current_case["parse_log"].append(f"[Info+] {text}")
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no

            elif current_field == "rep_fig":
                current_case["rep_fig_text"] += "\n" + text
                current_case["parse_log"].append(f"[RepFig+] {text}")

            elif current_field == "problem":
                current_case["problem"] += "\n" + text
                current_case["parse_log"].append(f"[Problem+] {text}")

            elif current_field == "spirit":
                current_case["spirit"] += "\n" + text
                current_case["parse_log"].append(f"[Spirit+] {text}")

            elif current_field == "key_point":
                current_case["key_point"] += "\n" + text
                current_case["parse_log"].append(f"[KeyPoint+] {text}")
            
            else:
                current_case["parse_log"].append(f"[Ignored] {text}")

        # 迴圈結束，存最後一筆
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
    st.info("👈 請上傳檔案。此版本包含詳細的歸類歷程記錄。")
else:
    st.subheader(f"📋 預覽")
    
    # === 新增：歷程檢視器 ===
    with st.expander("🕵️ 查看歸類歷程 (若資料消失請點我看原因)", expanded=False):
        for i, data in enumerate(st.session_state['slides_data']):
            st.markdown(f"**Case {i+1}: {data['raw_case_no']}**")
            st.json(data['parse_log']) # 直接顯示這筆案子吃到了什麼
    # ========================

    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.text(data['case_info'])
                
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    # 強制處理 None 或空字串
                    raw_text = data.get('rep_fig_text', "")
                    display_text = raw_text if raw_text and raw_text.strip() else "(Word中無代表圖資訊)"
                    st.warning(f"無圖片，將填入文字：\n{display_text[:50]}...")
                
                st.caption(f"重點：{data['key_point']}")

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 1. 左上：案號
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
                txBox = slide.shapes.add_textbox(img_left, img_top, img_width, img_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                
                raw_text = data.get('rep_fig_text', "")
                content_text = raw_text if raw_text and raw_text.strip() else "(Word中無代表圖資訊)"
                
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
            file_name="reconstructed_logic_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
