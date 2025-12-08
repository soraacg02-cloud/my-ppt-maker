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
st.set_page_config(page_title="PPT 重組生成器 (邏輯修復版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (邏輯修復版)")
st.caption("修正：解決因內文包含「公司/日期」等關鍵字，導致代表圖文字被截斷或消失的問題。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：依據關鍵字搜尋 PDF 並截圖 ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    if not target_fig_text:
        return None

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        # 為了搜尋精確，只取第一行非空文字
        lines = target_fig_text.split('\n')
        search_keyword = ""
        for line in lines:
            if line.strip():
                search_keyword = line.strip()
                break
        
        if not search_keyword:
            return None

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

# --- 函數：解析 Word 檔案 (嚴格狀態機版) ---
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
            
            # --- 1. 最高優先級：判斷是否為「新案件」的開始 (案號/索號) ---
            if "案號" in text or "索號" in text:
                # 只有遇到這兩個字，才百分之百確定是新的一案，或是該案的開頭
                
                # 如果已經有累積的資料，且不是正在寫同一個案號區塊，則存檔
                if current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                
                current_field = "case_info_block"
                # 這裡直接賦值，不使用 +=，因為這是一案的起點
                current_case["case_info"] = text 
                
                extracted_no = extract_patent_number_from_text(text)
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no
                
                debug_raw_lines.append(f"[Start Case] {text}")
                continue # 處理完就換下一行

            # --- 2. 判斷是否為其他「欄位標題」 ---
            
            if "解決問題" in text:
                current_field = "problem"
                content = re.sub(r'^[0-9.．]*\s*解決問題[:：]?\s*', '', text)
                current_case["problem"] = content
                debug_raw_lines.append(f"[Field: Problem] {text}")
                continue

            elif "發明精神" in text:
                current_field = "spirit"
                content = re.sub(r'^[0-9.．]*\s*發明精神[:：]?\s*', '', text)
                current_case["spirit"] = content
                debug_raw_lines.append(f"[Field: Spirit] {text}")
                continue

            elif "重點" in text:
                current_field = "key_point"
                content = re.sub(r'^[0-9.．]*\s*(一句)?重點[:：]?\s*', '', text)
                current_case["key_point"] = content
                debug_raw_lines.append(f"[Field: KeyPoint] {text}")
                continue

            elif "代表圖" in text:
                current_field = "rep_fig"
                content = re.sub(r'^[0-9.．]*\s*代表圖[:：]?\s*', '', text).strip()
                current_case["rep_fig_text"] = content
                debug_raw_lines.append(f"[Field: RepFig] {text}")
                continue

            # --- 3. 處理內容續行 (關鍵修正點) ---
            
            # 只有當目前還在 "case_info_block" (也就是左上角資訊區) 時，
            # 我們才把 "日期"、"申請日"、"公司" 當作資訊標題來處理。
            # 如果已經進入了 "代表圖" 或 "解決問題"，就算內文有 "公司"，也只是普通文字。
            
            is_header_keyword = any(k in text for k in ["日期", "申請日", "公司"])
            
            if current_field == "case_info_block":
                # 在資訊區塊，不管是不是關鍵字，都視為資訊的一部分
                current_case["case_info"] += "\n" + text
                # 隨時更新案號抓取
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no
                debug_raw_lines.append(f"  -> Add to CaseInfo: {text}")

            elif current_field == "rep_fig":
                # 在代表圖區塊，所有文字(包含換行、包含關鍵字)都屬於代表圖
                current_case["rep_fig_text"] += "\n" + text
                debug_raw_lines.append(f"  -> Add to RepFig: {text}")

            elif current_field == "problem":
                current_case["problem"] += "\n" + text
                debug_raw_lines.append(f"  -> Add to Problem: {text}")

            elif current_field == "spirit":
                current_case["spirit"] += "\n" + text
                debug_raw_lines.append(f"  -> Add to Spirit: {text}")

            elif current_field == "key_point":
                current_case["key_point"] += "\n" + text
                debug_raw_lines.append(f"  -> Add to KeyPoint: {text}")
                
            else:
                # 沒欄位歸屬的游離文字，暫時忽略或依需求處理
                debug_raw_lines.append(f"[Ignored] {text}")

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
        
        # Debug 資訊
        with st.expander("🔍 檢查 Word 解析邏輯 (Debug)", expanded=False):
            st.text("\n".join(raw_lines))
        
        # 讀取 PDF
        pdf_file_map = {}
        if pdf_files:
            for pdf in pdf_files:
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                pdf_file_map[clean_name] = pdf.read()

        # 配對
        match_count = 0
        
        with st.spinner("正在處理..."):
            for case in extracted_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"]
                
                matched_pdf_bytes = None
                
                for pdf_key, pdf_bytes in pdf_file_map.items():
                    if case_key and ((pdf_key.lower() in case_key.lower()) or (case_key.lower() in pdf_key.lower())):
                        if len(case_key) > 4: 
                            matched_pdf_bytes = pdf_bytes
                            break
                
                if matched_pdf_bytes and target_fig:
                    img_data = extract_specific_figure_from_pdf(matched_pdf_bytes, target_fig)
                    if img_data:
                        case["image_data"] = img_data
                        case["image_name"] = f"成功截取: {target_fig}"
                        match_count += 1
                    else:
                        case["image_name"] = f"找不到圖"
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
                st.text(data['case_info'])
                
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    # 顯示文字內容
                    display_text = data['rep_fig_text'] if data['rep_fig_text'].strip() else "(Word中無代表圖資訊)"
                    st.info(f"無圖片，將填入：\n{display_text}")
                
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
            file_name="fixed_logic_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
