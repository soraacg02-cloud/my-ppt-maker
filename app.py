import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from io import BytesIO
import docx
import fitz  # PyMuPDF，用來處理 PDF
from PIL import Image

# --- 設定網頁標題 ---
st.set_page_config(page_title="PPT 重組生成器 (進階版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (支援 Word+PDF 批次整合)")
st.caption("支援 Word 自動拆案，並可批次上傳 PDF 自動對應案號填入圖片。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：從 PDF 中提取圖片 (模擬抓取 Fig. 1) ---
def extract_image_from_pdf(pdf_stream):
    """
    從 PDF 檔案串流中提取圖片。
    策略：
    1. 優先搜尋含有 "Fig. 1", "Fig 1", "圖1" 文字的頁面。
    2. 若找不到文字，則回傳第一頁發現的圖片。
    """
    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        target_page_index = None
        
        # 1. 嘗試搜尋關鍵字所在的頁面
        for i, page in enumerate(doc):
            text = page.get_text()
            if "Fig. 1" in text or "Fig 1" in text or "圖1" in text or "圖 1" in text:
                target_page_index = i
                break
        
        # 如果沒找到關鍵字，預設從第一頁開始找
        if target_page_index is None:
            pages_to_check = range(len(doc))
        else:
            # 優先檢查找到的那一頁，之後檢查其他頁
            pages_to_check = [target_page_index] + [j for j in range(len(doc)) if j != target_page_index]

        for page_idx in pages_to_check:
            page = doc[page_idx]
            image_list = page.get_images(full=True)
            
            if image_list:
                # 找到圖片了，取出最大的一張 (避免抓到 icon 或 logo)
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # 簡單過濾過小的圖片 (例如小於 5KB 的可能是 logo)
                    if len(image_bytes) > 5120: 
                        return image_bytes

        return None
    except Exception as e:
        print(f"PDF 解析錯誤: {e}")
        return None

# --- 函數：從 PPT 中提取圖片與文字 (既有功能) ---
def extract_data_from_pptx(uploaded_pptx):
    try:
        prs = Presentation(uploaded_pptx)
        slide = prs.slides[0]
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

# --- 函數：解析 Word 檔案 (更新版：包含申請日) ---
def parse_word_file(uploaded_docx):
    """
    解析 Word，包含「案號」、「日期」、「申請日」合併為同一欄位。
    """
    try:
        doc = docx.Document(uploaded_docx)
        cases = []
        # 初始化
        current_case = {
            "case_info": "", 
            "problem": "", 
            "spirit": "", 
            "key_point": "", 
            "image_data": None, 
            "image_name": "Word匯入",
            "raw_case_no": "" # 用來做檔名比對的原始案號字串
        }
        current_field = None 

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue

            # --- 1. 案號 / 日期 / 申請日 (合併處理) ---
            if any(k in text for k in ["案號", "日期", "申請日"]):
                # 如果讀到「案號」，且目前已經有紀錄「案號」，代表進入下一案
                if "案號" in text and current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", 
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                
                current_field = "case_info_block"
                
                # 處理文字：保留標籤以便閱讀，但移除多餘空白
                # 如果是「案號」，順便存入 raw_case_no 供後續比對
                if "案號" in text:
                    clean_no = text.replace("案號", "").replace(":", "").replace("：", "").strip()
                    current_case["raw_case_no"] = clean_no
                
                # 將資訊串接到 case_info 欄位 (換行顯示)
                if current_case["case_info"]:
                    current_case["case_info"] += "\n" + text
                else:
                    current_case["case_info"] = text

            # --- 2. 解決問題 ---
            elif "解決問題" in text:
                current_field = "problem"
                clean_text = text.replace("解決問題", "").replace(":", "").replace("：", "").strip()
                current_case["problem"] = clean_text

            # --- 3. 發明精神 ---
            elif "發明精神" in text:
                current_field = "spirit"
                clean_text = text.replace("發明精神", "").replace(":", "").replace("：", "").strip()
                current_case["spirit"] = clean_text

            # --- 4. 一句重點 ---
            elif "重點" in text:
                current_field = "key_point"
                clean_text = text.replace("一句重點", "").replace("重點", "").replace(":", "").replace("：", "").strip()
                current_case["key_point"] = clean_text

            else:
                # 續行文字處理
                if current_field == "case_info_block":
                    current_case["case_info"] += " " + text
                elif current_field in ["problem", "spirit", "key_point"]:
                    current_case[current_field] += "\n" + text

        # 存入最後一筆
        if current_case["case_info"]:
            cases.append(current_case)
        
        return cases

    except Exception as e:
        st.error(f"解析 Word 時發生錯誤: {e}")
        return []

# --- 側邊欄 ---
with st.sidebar:
    st.header("1. 匯入資料來源")
    
    import_mode = st.radio("選擇匯入方式", ["Word + PDF 批次處理", "手動輸入 / PPT 提取"])

    if import_mode == "Word + PDF 批次處理":
        st.info("步驟 A：上傳 Word 檔 (含案號/日期/申請日/內文)")
        word_file = st.file_uploader("上傳 Word (.docx)", type=['docx'])
        
        st.info("步驟 B：上傳多個 PDF 檔 (檔名需包含案號)")
        pdf_files = st.file_uploader("上傳 PDF (.pdf)", type=['pdf'], accept_multiple_files=True)
        
        if word_file and st.button("🔄 開始批次整合", type="primary"):
            # 1. 解析 Word
            extracted_cases = parse_word_file(word_file)
            
            # 2. 預處理 PDF 圖片
            pdf_images = {} # 格式: {'檔名關鍵字': image_bytes}
            if pdf_files:
                with st.spinner("正在分析 PDF 圖片..."):
                    for pdf in pdf_files:
                        # 去除副檔名，轉小寫以利比對
                        clean_name = pdf.name.rsplit('.', 1)[0].lower()
                        img_data = extract_image_from_pdf(pdf.read())
                        if img_data:
                            pdf_images[clean_name] = img_data
            
            # 3. 進行配對
            match_count = 0
            for case in extracted_cases:
                # 取得 Word 中的案號 (轉小寫去除空白)
                case_key = case["raw_case_no"].lower().replace(" ", "")
                
                # 嘗試比對 PDF 檔名
                # 邏輯：檢查 PDF 檔名是否包含在案號中，或案號是否包含在 PDF 檔名中
                matched_img = None
                matched_name = ""
                
                for pdf_name, img_bytes in pdf_images.items():
                    # 清理 pdf 名稱
                    clean_pdf_name = pdf_name.replace(" ", "")
                    
                    # 寬鬆比對：只要有一方包含另一方就算對應
                    if (clean_pdf_name in case_key and len(clean_pdf_name) > 3) or \
                       (case_key in clean_pdf_name and len(case_key) > 3):
                        matched_img = img_bytes
                        matched_name = pdf_name
                        break
                
                if matched_img:
                    case["image_data"] = matched_img
                    case["image_name"] = f"PDF: {matched_name}"
                    match_count += 1
                else:
                    case["image_name"] = "無對應 PDF"

            if extracted_cases:
                st.session_state['slides_data'].extend(extracted_cases)
                st.success(f"匯入成功！共 {len(extracted_cases)} 筆資料，其中 {match_count} 筆成功配對圖片。")
            else:
                st.warning("Word 解析失敗或無資料。")

    else:
        # --- 手動 / PPT 模式 ---
        uploaded_file = st.file_uploader("上傳 PPT (.pptx) 提取圖文", type=['pptx'])
        ppt_image_blob = None
        extracted_txt_content = ""

        if uploaded_file:
            ppt_image_blob, extracted_txt_content = extract_data_from_pptx(uploaded_file)
            if ppt_image_blob:
                st.image(ppt_image_blob, caption="PPT 圖片", use_column_width=True)
            with st.expander("PPT 文字"):
                st.text_area("內容", extracted_txt_content)

        st.divider()
        st.header("編輯內容")
        # 這裡修改提示文字，讓使用者知道可以輸入申請日
        case_info = st.text_input("1. 案號 / 日期 / 申請日")
        uploaded_img = st.file_uploader("2. 上傳圖片 (選填)", type=['png', 'jpg'])
        problem = st.text_area("3. 解決問題")
        spirit = st.text_area("4. 發明精神")
        key_point = st.text_input("5. 一句重點")

        if st.button("➕ 加入此頁"):
            img_data = ppt_image_blob
            img_name = "PPT提取"
            if uploaded_img:
                img_data = uploaded_img.getvalue()
                img_name = uploaded_img.name
            
            st.session_state['slides_data'].append({
                "case_info": case_info,
                "problem": problem,
                "spirit": spirit,
                "key_point": key_point,
                "image_data": img_data,
                "image_name": img_name
            })
            st.success("已新增！")

    if st.session_state['slides_data']:
        st.divider()
        if st.button("🗑️ 清除所有"):
            st.session_state['slides_data'] = []
            st.rerun()

# --- 主畫面 ---
if not st.session_state['slides_data']:
    st.info("👈 請從左側匯入資料。")
else:
    st.subheader(f"📋 預覽 ({len(st.session_state['slides_data'])} 頁)")
    cols = st.columns(3)
    for i, data in enumerate(st.session_state['slides_data']):
        with cols[i % 3]:
            with st.container(border=True):
                st.markdown(f"**第 {i+1} 頁**")
                # 這裡會顯示包含申請日的多行文字
                st.text(data['case_info'])
                if data['image_data']:
                    st.image(data['image_data'], caption=data.get('image_name', ''), use_column_width=True)
                else:
                    st.markdown("*(無圖片)*")
                st.caption(f"重點：{data['key_point']}")

    st.divider()

    # --- PPT 生成邏輯 ---
    def generate_ppt(slides_data):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for data in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 1. 左上角：案號 / 日期 / 申請日
            left, top, width, height = Inches(0.5), Inches(0.5), Inches(5.0), Inches(1.5) # 高度增加以容納多行
            txBox = slide.shapes.add_textbox(left, top, width, height)
            p = txBox.text_frame.add_paragraph()
            p.text = data['case_info'] # 這裡會直接填入包含申請日的完整字串
            p.font.size = Pt(20) # 字體稍微調小一點以適應多行
            p.font.bold = True
            
            # 2. 右上角：圖片 (綠框位置)
            if data['image_data']:
                image_stream = BytesIO(data['image_data'])
                # 限制高度 4 英吋，位置固定右上
                slide.shapes.add_picture(image_stream, Inches(5.5), Inches(0.5), height=Inches(4.0))

            # 3. 中下：解決問題 / 發明精神
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
            label="📥 下載結果",
            data=binary_output,
            file_name="final_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
