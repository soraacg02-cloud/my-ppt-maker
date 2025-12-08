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
st.set_page_config(page_title="PPT 重組生成器 (強效版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (強效截圖版)")
st.caption("升級版：使用「頁面截圖」技術，解決專利線條圖無法提取的問題。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：強效 PDF 截圖 (Render Page) ---
def extract_image_from_pdf_robust(pdf_stream):
    """
    使用渲染技術將 PDF 頁面轉為圖片。
    策略：
    1. 搜尋含有 "Fig. 1", "圖 1", "圖1" 的頁面。
    2. 若找不到，預設抓取「第一頁」(通常是摘要頁，含代表圖)。
    3. 將該頁面「截圖」存為圖片檔。
    """
    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        target_page_index = 0 # 預設第一頁
        found_keyword = False
        
        # 1. 嘗試搜尋關鍵字所在的頁面
        for i, page in enumerate(doc):
            text = page.get_text()
            # 搜尋常見的圖式標記
            if any(k in text for k in ["Fig. 1", "Fig 1", "FIG. 1", "圖 1", "圖1", "代表圖"]):
                target_page_index = i
                found_keyword = True
                break
        
        # 2. 將目標頁面轉為圖片 (截圖)
        page = doc[target_page_index]
        # 設定解析度 (Matrix(2, 2) 代表放大 2 倍，讓圖片更清晰)
        zoom = 2 
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        # 轉為 PNG 二進位資料
        return pix.tobytes("png")

    except Exception as e:
        print(f"PDF 解析錯誤: {e}")
        return None

# --- 函數：從 Word 中提取資料 ---
def parse_word_file(uploaded_docx):
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        current_case = {
            "case_info": "", "problem": "", "spirit": "", "key_point": "", 
            "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
        }
        current_field = None 

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            # 關鍵字判斷
            if any(k in text for k in ["案號", "日期", "申請日"]):
                if "案號" in text and current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", 
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                current_field = "case_info_block"
                
                # 抓取原始案號用於比對 (移除標點符號，只留英數字)
                if "案號" in text:
                    raw_no = text.split("：")[-1] if "：" in text else text.split(":")[-1]
                    # 只保留英數字以便比對 (去除空白、斜線等)
                    clean_no = re.sub(r'[^a-zA-Z0-9]', '', raw_no)
                    current_case["raw_case_no"] = clean_no
                
                current_case["case_info"] += text + "\n"

            elif "解決問題" in text:
                current_field = "problem"
                current_case["problem"] = text.replace("解決問題", "").replace(":", "").replace("：", "").strip()

            elif "發明精神" in text:
                current_field = "spirit"
                current_case["spirit"] = text.replace("發明精神", "").replace(":", "").replace("：", "").strip()

            elif "重點" in text:
                current_field = "key_point"
                current_case["key_point"] = text.replace("一句重點", "").replace("重點", "").replace(":", "").replace("：", "").strip()

            else:
                if current_field == "case_info_block":
                    current_case["case_info"] += text + " "
                elif current_field in ["problem", "spirit", "key_point"]:
                    current_case[current_field] += "\n" + text

        if current_case["case_info"]:
            cases.append(current_case)
        return cases
    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料")
    st.info("步驟 A：上傳 Word (.docx)")
    word_file = st.file_uploader("Word 檔案", type=['docx'])
    
    st.info("步驟 B：上傳多個 PDF (.pdf)")
    pdf_files = st.file_uploader("PDF 檔案 (可多選)", type=['pdf'], accept_multiple_files=True)
    
    if word_file and st.button("🔄 開始強效整合", type="primary"):
        # 1. 解析 Word
        extracted_cases = parse_word_file(word_file)
        
        # 2. 處理 PDF (轉圖片)
        pdf_images = {}
        pdf_debug_names = [] # 用來除錯顯示
        if pdf_files:
            with st.spinner("正在將 PDF 頁面轉為圖片..."):
                for pdf in pdf_files:
                    try:
                        # 檔名清理：只留英數字
                        clean_name = re.sub(r'[^a-zA-Z0-9]', '', pdf.name.rsplit('.', 1)[0])
                        pdf_debug_names.append(f"{pdf.name} -> 識別為: {clean_name}")
                        
                        img_data = extract_image_from_pdf_robust(pdf.read())
                        if img_data:
                            pdf_images[clean_name] = img_data
                    except Exception as e:
                        st.error(f"處理 PDF {pdf.name} 時失敗: {e}")

        # 3. 進行配對
        match_count = 0
        debug_logs = []
        
        for case in extracted_cases:
            case_key = case["raw_case_no"] # 這是從 Word 抓出來並清理過的案號
            matched_img = None
            matched_name = ""
            
            # 比對邏輯：檢查「PDF 檔名」是否包含「案號」，反之亦然
            for pdf_key, img_bytes in pdf_images.items():
                # 轉小寫比對
                if (pdf_key.lower() in case_key.lower() and len(pdf_key) > 3) or \
                   (case_key.lower() in pdf_key.lower() and len(case_key) > 3):
                    matched_img = img_bytes
                    matched_name = pdf_key
                    break
            
            debug_logs.append(f"Word案號: {case_key} | 配對結果: {matched_name if matched_name else '失敗'}")

            if matched_img:
                case["image_data"] = matched_img
                case["image_name"] = f"PDF: {matched_name}"
                match_count += 1
            else:
                case["image_name"] = "無圖片"

        # 存入 Session
        if extracted_cases:
            st.session_state['slides_data'].extend(extracted_cases)
            st.success(f"匯入 {len(extracted_cases)} 筆，成功配對 {match_count} 張圖片！")
            
            # --- 顯示診斷資訊 (幫助您除錯) ---
            with st.expander("🕵️ 查看配對診斷報告 (如果圖片沒出來請看這)", expanded=False):
                st.write("### 1. 系統讀到的 PDF 檔名")
                st.write(pdf_debug_names)
                st.write("### 2. Word 與 PDF 配對詳情")
                st.write(debug_logs)
        else:
            st.warning("Word 解析無資料。")

    # 清除按鈕
    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面：預覽與生成 ---
if not st.session_state['slides_data']:
    st.info("👈 請從左側開始匯入資料。本版本支援專利線條圖提取。")
else:
    st.subheader(f"📋 預覽 ({len(st.session_state['slides_data'])} 頁)")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                st.text(data['case_info'].strip())
                if data['image_data']:
                    st.image(data['image_data'], caption=data.get('image_name', ''), use_column_width=True)
                else:
                    st.warning("❌ 無圖片")
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
            
            # 2. 右上：圖片 (綠框)
            if data['image_data']:
                image_stream = BytesIO(data['image_data'])
                # 因為是截圖，可能包含整頁白邊，這裡設定高度限制，讓它自動縮放
                slide.shapes.add_picture(image_stream, Inches(5.5), Inches(0.5), height=Inches(4.0))

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
            file_name="patent_slides_robust.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
