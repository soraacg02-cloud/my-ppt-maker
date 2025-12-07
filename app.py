import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

st.set_page_config(page_title="PPT產生器", layout="centered")
st.title("PPT 自動排版神器")

# 1. 設定頁數
num = st.number_input("你要做幾頁？", min_value=1, value=3)
slides_data = []

# 2. 顯示輸入框
for i in range(int(num)):
    st.markdown(f"### 第 {i+1} 頁")
    txt = st.text_input(f"文字 (第 {i+1} 頁)", key=f"t{i}")
    img = st.file_uploader(f"圖片 (第 {i+1} 頁)", type=["png","jpg"], key=f"i{i}")
    slides_data.append({"text": txt, "image": img})

# 3. 按鈕生成
if st.button("下載 PPT"):
    prs = Presentation()
    prs.slide_width = Inches(13.333) # 16:9 寬螢幕
    prs.slide_height = Inches(7.5)
    
    has_data = False
    for d in slides_data:
        if not d["text"] and not d["image"]: continue
        has_data = True
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # 空白頁
        
        # 貼圖 (左邊)
        if d["image"]:
            slide.shapes.add_picture(d["image"], Inches(1), Inches(1.5), width=Inches(6))
            
        # 貼字 (右邊)
        tb = slide.shapes.add_textbox(Inches(7.5), Inches(2), Inches(5), Inches(3))
        tf = tb.text_frame
        tf.text = d["text"]
        for p in tf.paragraphs:
            p.font.size = Pt(24)
            p.font.bold = True

    if has_data:
        out = BytesIO()
        prs.save(out)
        out.seek(0)
        st.download_button("點我下載檔案", out, "mypresentation.pptx")
    else:
        st.warning("請至少輸入一點內容！")
