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
st.set_page_config(page_title="PPT 重組生成器 (精準修正版)", page_icon="📑", layout="wide")
st.title("📑 PPT 重組生成器 (代表圖文字修正版)")
st.caption("修正：解決代表圖包含數字時會被誤刪的問題，並新增原始資料檢查功能。")

# --- 初始化 Session State ---
if 'slides_data' not in st.session_state:
    st.session_state['slides_data'] = []

# --- 函數：依據關鍵字搜尋 PDF 並截圖 ---
def extract_specific_figure_from_pdf(pdf_stream, target_fig_text):
    if not target_fig_text:
        return None

    try:
        doc = fitz.open(stream=pdf_stream, filetype="pdf")
        # 只取第一行搜尋
        search_keyword = target_fig_text.split('\n')[0].strip()
        # 移除空白以增加比對成功率
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

# --- 函數：解析 Word 檔案 (修正代表圖抓取邏輯) ---
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
        
        # 用來Debug用的原始紀錄
        debug_raw_lines = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue
            
            debug_raw_lines.append(text) # 紀錄原始文字供檢查

            # --- 關鍵字判斷 ---
            
            # 1. 案號 / 日期 / 公司
            if any(k in text for k in ["案號", "日期", "申請日", "索號", "公司"]):
                if ("案號" in text or "索號" in text) and current_case["case_info"] and current_field != "case_info_block":
                    cases.append(current_case)
                    current_case = {
                        "case_info": "", "problem": "", "spirit": "", "key_point": "", "rep_fig_text": "",
                        "image_data": None, "image_name": "Word匯入", "raw_case_no": ""
                    }
                current_field = "case_info_block"
                
                if current_case["case_info"]:
                    current_case["case_info"] += "\n" + text
                else:
                    current_case["case_info"] = text
                
                extracted_no = extract_patent_number_from_text(current_case["case_info"])
                if extracted_no:
                    current_case["raw_case_no"] = extracted_no

            # 2. 解決問題
            elif "解決問題" in text:
                current_field = "problem"
                # 使用 Regex 移除標題，避免誤刪內容
                content = re.sub(r'^[0-9.．]*\s*解決問題[:：]?\s*', '', text)
                current_case["problem"] = content

            # 3. 發明精神
            elif "發明精神" in text:
                current_field = "spirit"
                content = re.sub(r'^[0-9.．]*\s*發明精神[:：]?\s*', '', text)
                current_case["spirit"] = content

            # 4. 一句重點
            elif "重點" in text:
                current_field = "key_point"
                content = re.sub(r'^[0-9.．]*\s*(一句)?重點[:：]?\s*', '', text)
                current_case["key_point"] = content

            # 5. 代表圖 (修正重點)
            elif "代表圖" in text:
                current_field = "rep_fig"
                # 舊邏輯: text.replace("5", "") -> 錯誤！會把內容的 5 刪掉
                # 新邏輯: 使用 Regex 只移除「開頭的編號」和「代表圖」標籤
                # 說明: ^[0-9.．]* 匹配開頭的數字和點, \s*代表圖[:：]? 匹配代表圖和冒號
                content = re.sub(r'^[0-9.．]*\s*代表圖[:：]?\s*', '', text).strip()
                current_case["rep_fig_text"] = content

            else:
                # 續行文字處理
                if current_field == "case_info_block":
                    current_case["case_info"] += "\n" + text
                    extracted_no = extract_patent_number_from_text(current_case["case_info"])
                    if extracted_no:
                        current_case["raw_case_no"] = extracted_no
                elif current_field in ["problem", "spirit", "key_point"]:
                    current_case[current_field] += "\n" + text
                elif current_field == "rep_fig":
                    current_case["rep_fig_text"] += "\n" + text 

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
        
        # 顯示原始資料檢查器 (Debug用)
        with st.expander("🔍 檢查 Word 讀取到的內容 (若有問題請看這)", expanded=
