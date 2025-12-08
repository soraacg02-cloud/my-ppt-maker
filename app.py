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
import pandas as pd

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器 (簡潔版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (完整版)")
st.caption("支援多檔上傳、自動排序 (申請人 -> 日期)、表格讀取與錯誤診斷。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []
if 'status_report' not in st.session_state:
    st.session_state['status_report'] = []

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
        return None, "Word 中未指定代表圖文字"

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
            return None, "無法從說明文字中識別出圖號"

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
            return pix.tobytes("png"), f"成功"
            
        return None, f"PDF 中找不到關鍵字「{target_keyword}」"
    except Exception as e:
        return None, f"PDF 解析發生錯誤: {str(e)}"

# --- 函數：提取專利號 ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match:
        return match.group(1)
    return ""

# --- 函數：提取日期 (用於排序) ---
def extract_date_for_sort(text):
    match = re.search(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', text)
    if match:
        return f"{match.group(1)}{match.group(2).zfill(2)}{match.group(3).zfill(2)}"
    return "99999999"

# --- 函數：提取公司/申請人 (用於排序) ---
def extract_company_for_sort(text):
    lines = text.split('\n')
    for line in lines:
        if "公司" in line or "申請人" in line:
            if "案號" in line and "日期" in line: # 跳過標題行
                continue
            return line.replace("公司", "").replace("申請人", "").replace("：", "").replace(":", "").strip()
    return "ZZZ"

# --- 函數：解析 Word 檔案 ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
            "image_data": None, "image_name": "Word匯入", "raw_case_no": "",
            "sort_date": "99999999", "sort_company": "ZZZ",
            "source_file": uploaded_docx.name,
            "missing_fields": []
        }
        current_field = None 
        
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
        
        for text in all_lines:
            # A. 新案件判斷
            if "案號" in text or "索號" in text:
                if current_case["case_info"] and current_field != "case_info_block":
                    if not current_case["problem"]: current_case["missing_fields"].append("解決問題")
                    if not current_case["spirit"]: current_case["missing_fields"].append("發明精神")
                    if not current_case["key_point"]: current_case["missing_fields"].append("一句重點")
                    cases.append(current_case)
                    
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": "",
                        "sort_date": "99999999", "sort_company": "ZZZ",
                        "source_file": uploaded_docx.name,
                        "missing_fields": []
                    }
                
                current_field = "case_info_block"
                current_case["case_info"] = text
                
                extracted_no = extract_patent_number_from_text(text)
                if extracted_no: current_case["raw_case_no"] = extracted_no
                current_case["sort_date"] = extract_date_for_sort(text)
                current_case["sort_company"] = extract_company_for_sort(text)
                continue

            # B. 欄位切換
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

            # C. 內容填充
            if current_field == "case_info_block":
                current_case["case_info"] += "\n" + text
                if current_case["sort_date"] == "99999999":
                    current_case["sort_date"] = extract_date_for_sort(text)
                
                extracted_comp = extract_company_for_sort(current_case["case_info"])
                if extracted_comp != "ZZZ":
                    current_case["sort_company"] = extracted_comp
                
                if not current_case["raw_case_no"]:
                    extracted_no = extract_patent_number_from_text(text)
                    if extracted_no: current_case["raw_case_no"] = extracted_no

            elif current_field == "rep_fig":
                current_case["rep_fig_text"] += "\n" + text
            elif current_field == "problem":
                current_case["problem"] += "\n" + text
            elif current_field == "spirit":
                current_case["spirit"] += "\n" + text
            elif current_field == "key_point":
                current_case["key_point"] += "\n" + text

        # 存最後一筆
        if current_case["case_info"]:
            if not current_case["problem"]: current_case["missing_fields"].append("解決問題")
            if not current_case["spirit"]: current_case["missing_fields"].append("發明精神")
            if not current_case["key_point"]: current_case["missing_fields"].append("一句重點")
            cases.append(current_case)
            
        return cases

    except Exception as e:
        st.error(f"解析 Word 時發生錯誤 ({uploaded_docx.name}): {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    word_files = st.file_uploader("Word 檔案 (可多選)", type=['docx'], accept_multiple_files=True)
    pdf_files = st.file_uploader("PDF 檔案 (可多選)", type=['pdf'], accept_multiple_files=True)
    
    if word_files and st.button("🔄 開始智能整合", type="primary"):
        all_cases = []
        status_report_list = []
        
        # 1. 處理 Word
        for word_file in word_files:
            cases = parse_word_file(word_file)
            all_cases.extend(cases)
        
        # 2. 準備 PDF
        pdf_file_map = {}
        if pdf_files:
            for pdf in pdf_files:
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                pdf_file_map[clean_name] = pdf.read()

        match_count = 0
        
        with st.spinner("正在處理..."):
            for case in all_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"]
                
                status = {
                    "來源檔案": case["source_file"],
                    "案號": case_key if case_key else "(無法辨識)",
                    "申請人/公司": case["sort_company"] if case["sort_company"] != "ZZZ" else "(未找到)",
                    "日期": case["sort_date"] if case["sort_date"] != "99999999" else "(未找到)",
                    "圖片狀態": "未處理",
                    "錯誤原因": "",
                    "缺漏欄位": ", ".join(case["missing_fields"]) if case["missing_fields"] else "無"
                }

                matched_pdf_bytes = None
                
                for pdf_key, pdf_bytes in pdf_file_map.items():
                    if case_key and ((pdf_key.lower() in case_key.lower()) or (case_key.lower() in pdf_key.lower())):
                        if len(case_key) > 4: 
                            matched_pdf_bytes = pdf_bytes
                            break
                
                if matched_pdf_bytes:
                    img_data, msg = extract_specific_figure_from_pdf(matched_pdf_bytes, target_fig)
                    if img_data:
                        case["image_data"] = img_data
                        case["image_name"] = f"成功"
                        status["圖片狀態"] = "✅ 成功"
                        match_count += 1
                    else:
                        case["image_name"] = "缺圖"
                        status["圖片狀態"] = "⚠️ 缺圖"
                        status["錯誤原因"] = msg
                else:
                    if not target_fig:
                        status["圖片狀態"] = "⚠️ 缺資訊"
                        status["錯誤原因"] = "Word 中未填寫「代表圖」欄位"
                    else:
                        status["圖片狀態"] = "❌ 找不到 PDF"
                        status["錯誤原因"] = f"找不到檔名包含「{case_key}」的 PDF"
                
                status_report_list.append(status)

        # 3. 排序 (公司/申請人 A-Z -> 日期早到晚)
        all_cases.sort(key=lambda x: (x["sort_company"].upper(), x["sort_date"]))
        status_report_list.sort(key=lambda x: (x["申請人/公司"].upper(), x["日期"]))

        if all_cases:
            st.session_state['slides_data'] = all_cases
            st.session_state['status_report'] = status_report_list
            st.success(f"處理完成！共 {len(all_cases)} 筆，成功截取 {match_count} 張圖。")
        else:
            st.warning("Word 解析無資料。")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.session_state['status_report'] = []
            st.rerun()

# --- 主畫面 ---
if not st.session_state['slides_data']:
    st.info("👈 請先在左側上傳檔案。")
else:
    # 1. 簡報預覽
    st.subheader(f"📋 簡報預覽 (已排序)")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.caption(f"{data['sort_company']} | {data['sort_date']}")
                st.text(data['case_info'][:100] + "...")
                
                if data['image_data']:
                    st.image(data['image_data'], use_column_width=True)
                else:
                    raw_text = data.get('rep_fig_text', "")
                    display_text = raw_text if raw_text and raw_text.strip() else "(Word中無代表圖資訊)"
                    st.warning(f"無圖片，將填入文字：\n{display_text[:50]}...")
                st.caption(f"重點：{data['key_point']}")

    # PPT 下載
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 左上
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
                    p.alignment = PP_ALIGN.LEFT

            # 右上
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
                        p.alignment = PP_ALIGN.LEFT

            # 中下 & 底部
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
            shape.text_frame.vertical_anchor = MSO_SHAPE.RECTANGLE
        return prs

    st.divider()
    if st.button("🚀 生成 PowerPoint (.pptx)", type="primary"):
        prs = generate_ppt(st.session_state['slides_data'])
        binary_output = BytesIO()
        prs.save(binary_output)
        binary_output.seek(0)
        st.download_button("📥 下載 PPT", binary_output, "final_slides.pptx")

    # 2. 診斷表格 (最下面)
    st.divider()
    st.subheader("📊 處理結果診斷報告")
    if st.session_state['status_report']:
        df = pd.DataFrame(st.session_state['status_report'])
        st.dataframe(df, hide_index=True)
