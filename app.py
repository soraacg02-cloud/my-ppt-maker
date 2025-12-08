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
st.set_page_config(page_title="PPT 重組生成器 (指定圖式版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (指定代表圖版)")
st.caption("依據 Word 指定的「代表圖」名稱，自動從 PDF 截取對應頁面；若截取失敗則填入文字。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：依據關鍵字搜尋 PDF 並截圖 ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    """
    在 PDF 中搜尋 target_fig_text (例如 "圖1")。
    若找到包含該文字的頁面，則將該頁截圖回傳。
    """
    if not target_fig_text:
        return None

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        
        # 預處理：移除目標文字的空白，提高比對成功率 (例如 "圖 1" -> "圖1")
        clean_target = target_fig_text.replace(" ", "").strip()
        
        found_page_index = None

        # 遍歷每一頁找文字
        for i, page in enumerate(doc):
            page_text = page.get_text()
            # 移除頁面文字的空白來比對
            clean_page_text = page_text.replace(" ", "")
            
            if clean_target in clean_page_text:
                found_page_index = i
                break
        
        # 如果找到了，進行截圖
        if found_page_index is not None:
            page = doc[found_page_index]
            mat = fitz.Matrix(2, 2) # 放大 2 倍清晰度
            pix = page.get_pixmap(matrix=mat)
            return pix.tobytes("png")
            
        return None # 沒找到對應文字

    except Exception as e:
        print(f"PDF 解析錯誤: {e}")
        return None

# --- 函數：從文字中提取專利號 (用於檔名配對) ---
def extract_patent_number_from_text(text):
    clean_text = text.replace("：", ":").replace(" ", "")
    match = re.search(r'([a-zA-Z]{2,4}\d+[a-zA-Z]?)', clean_text)
    if match:
        return match.group(1)
    return ""

# --- 函數：解析 Word 檔案 (新增：5.代表圖) ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        # 初始化結構，新增 rep_fig_text
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
            "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
        }
        current_field = None 

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            # --- 關鍵字判斷 ---
            
            # 1. 案號 / 日期
            if any(k in text for k in ["案號", "日期", "申請日", "索號"]):
                # 遇到新案號，先存上一筆
                if ("案號" in text or "索號" in text) and current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                current_field = "case_info_block"
                current_case["case_info"] += text + "\n"
                
                # 嘗試提取案號 (CN/TW...)
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no

            # 2. 解決問題
            elif "解決問題" in text:
                current_field = "problem"
                current_case["problem"] = text.replace("解決問題", "").replace(":", "").replace("：", "").strip()

            # 3. 發明精神
            elif "發明精神" in text:
                current_field = "spirit"
                current_case["spirit"] = text.replace("發明精神", "").replace(":", "").replace("：", "").strip()

            # 4. 一句重點
            elif "重點" in text:
                current_field = "key_point"
                current_case["key_point"] = text.replace("一句重點", "").replace("重點", "").replace(":", "").replace("：", "").strip()

            # 5. 代表圖 (新增功能)
            elif "代表圖" in text:
                current_field = "rep_fig"
                # 清理文字，只留下 "圖1" 或 "Fig. 2" 這種內容
                clean_fig = text.replace("5", "").replace(".", "").replace("代表圖", "").replace(":", "").replace("：", "").strip()
                current_case["rep_fig_text"] = clean_fig

            else:
                # 續行文字處理
                if current_field == "case_info_block":
                    current_case["case_info"] += text + "\n"
                    # 若續行包含案號，再次嘗試提取
                    extracted_no = extract_patent_number_from_text(current_case["case_info"])
                    if extracted_no:
                        current_case["raw_case_no"] = extracted_no
                elif current_field in ["problem", "spirit", "key_point"]:
                    current_case[current_field] += "\n" + text
                elif current_field == "rep_fig":
                    current_case["rep_fig_text"] += text # 代表圖若有換行也接上去

        if current_case["case_info"]:
            cases.append(current_case)
        return cases
    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    st.info("請上傳包含「5. 代表圖」欄位的 Word 檔。")
    word_file = st.file_uploader("Word 檔案 (.docx)", type=['docx'])
    pdf_files = st.file_uploader("PDF 檔案 (.pdf)", type=['pdf'], accept_multiple_files=True)
    
    if word_file and st.button("🔄 開始智能整合", type="primary"):
        # 1. 解析 Word
        extracted_cases = parse_word_file(word_file)
        
        # 2. 讀取 PDF (暫存於記憶體，不預先轉圖，改為按需搜尋)
        pdf_file_map = {} # 格式: {'clean_filename': pdf_bytes}
        pdf_debug_names = []
        
        if pdf_files:
            for pdf in pdf_files:
                clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                pdf_file_map[clean_name] = pdf.read() # 讀取二進位資料
                pdf_debug_names.append(f"{pdf.name} -> {clean_name}")

        # 3. 進行配對與抓圖
        match_count = 0
        debug_logs = []
        
        with st.spinner("正在搜尋指定的代表圖..."):
            for case in extracted_cases:
                case_key = case["raw_case_no"]
                target_fig = case["rep_fig_text"] # 例如 "圖1"
                
                matched_pdf_bytes = None
                matched_name = ""
                
                # 尋找對應的 PDF
                for pdf_key, pdf_bytes in pdf_file_map.items():
                    if case_key and ((pdf_key.lower() in case_key.lower()) or (case_key.lower() in pdf_key.lower())):
                        if len(case_key) > 4: 
                            matched_pdf_bytes = pdf_bytes
                            matched_name = pdf_key
                            break
                
                # 若找到 PDF，則去 PDF 裡找代表圖
                if matched_pdf_bytes and target_fig:
                    img_data = extract_specific_figure_from_pdf(matched_pdf_bytes, target_fig)
                    if img_data:
                        case["image_data"] = img_data
                        case["image_name"] = f"成功截取: {target_fig}"
                        match_count += 1
                    else:
                        case["image_name"] = f"找不到「{target_fig}」"
                        debug_logs.append(f"案號 {case_key}: 找到PDF但找不到 '{target_fig}'，將使用文字替代。")
                else:
                    if not matched_pdf_bytes:
                        debug_logs.append(f"案號 {case_key}: 找不到對應 PDF。")
                    if not target_fig:
                        debug_logs.append(f"案號 {case_key}: Word 中未指定代表圖。")

        if extracted_cases:
            st.session_state['slides_data'].extend(extracted_cases)
            st.success(f"匯入 {len(extracted_cases)} 筆，圖片截取成功 {match_count} 筆！")
            if debug_logs:
                with st.expander("查看處理詳情", expanded=False):
                    st.write(debug_logs)
        else:
            st.warning("Word 解析無資料。")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：預覽與生成 ---
if not st.session_state['slides_data']:
    st.info("👈 請上傳 Word 與 PDF。程式將依據 Word 內的「5. 代表圖」去 PDF 抓圖。")
else:
    st.subheader(f"📋 預覽 ({len(st.session_state['slides_data'])} 頁)")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.caption(f"識別號: {data['raw_case_no']}")
                
                # 預覽區顯示邏輯
                if data['image_data']:
                    st.image(data['image_data'], caption=data.get('image_name', ''), use_column_width=True)
                else:
                    # 如果沒圖片，顯示將會填入的替代文字
                    st.warning(f"❌ 無截圖，將填入文字：\n\n「{data['rep_fig_text']}」")
                
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
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(1.5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
            p.text = data['case_info']
            p.font.size = Pt(20)
            p.font.bold = True
            
            # 2. 右上：圖片或替代文字 (綠框位置)
            # 位置定義
            img_left = Inches(5.5)
            img_top = Inches(0.5)
            img_height = Inches(4.0)
            img_width = Inches(7.0) # 給文字框用的寬度

            if data['image_data']:
                # === 情況 A: 有抓到圖 ===
                image_stream = BytesIO(data['image_data'])
                slide.shapes.add_picture(image_stream, img_left, img_top, height=img_height)
            else:
                # === 情況 B: 沒圖，填入 Word 指定的文字 ===
                # 建立一個文字方塊在原本放圖的位置
                txBox = slide.shapes.add_textbox(img_left, img_top, img_width, img_height)
                tf = txBox.text_frame
                tf.word_wrap = True
                
                p = tf.add_paragraph()
                p.text = data['rep_fig_text'] if data['rep_fig_text'] else "(未指定代表圖)"
                p.font.size = Pt(40) # 字體大一點，置中顯示
                p.font.bold = True
                p.font.color.rgb = RGBColor(128, 128, 128) # 灰色文字
                p.alignment = PP_ALIGN.CENTER
                
                # 垂直置中 (利用 textbox 的屬性)
                txBox.text_frame.vertical_anchor = MSO_SHAPE.RECTANGLE # 設為垂直置中效果

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
            file_name="specified_figure_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
