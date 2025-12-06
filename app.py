import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from io import BytesIO
import os

# ==========================================
# 1. 核心邏輯：提取內容
# ==========================================
def extract_content_from_pptx(file_obj):
    prs = Presentation(file_obj)
    extracted_data = []

    for slide in prs.slides:
        slide_content = {"title": "", "content": ""}
        
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            slide_content["title"] = slide.shapes.title.text_frame.text
        
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1 and shape.has_text_frame:
                slide_content["content"] = shape.text_frame.text
                break 
        
        if slide_content["title"] or slide_content["content"]:
            extracted_data.append(slide_content)
    return extracted_data

# ==========================================
# 2. 進階樣式生成邏輯 (包含繪圖與排版)
# ==========================================
def create_styled_pptx(data, style_mode):
    prs = Presentation()
    # 設定投影片大小為寬螢幕 16:9 (13.33 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for item in data:
        # 使用空白版型 (Layout 6)，讓我們完全控制位置
        slide_layout = prs.slide_layouts[6] 
        slide = prs.slides.add_slide(slide_layout)
        shapes = slide.shapes

        # -----------------------------------------------
        # 風格 A: 現代側邊欄 (Modern Sidebar)
        # 特色：左側 1/4 是深色色塊，右側是內容
        # -----------------------------------------------
        if style_mode == "sidebar":
            # 1. 畫左側色塊 (深藍)
            sidebar = shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 
                Inches(0), Inches(0), Inches(3.5), Inches(7.5) # x, y, w, h
            )
            sidebar.fill.solid()
            sidebar.fill.fore_color.rgb = RGBColor(44, 62, 80)
            sidebar.line.fill.background() # 去除邊框

            # 2. 新增標題 (在右側白底區)
            title_box = shapes.add_textbox(Inches(4), Inches(0.5), Inches(8.5), Inches(1.5))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = RGBColor(44, 62, 80)
            p.font.name = "Arial Black"

            # 3. 新增內文 (在右側)
            content_box = shapes.add_textbox(Inches(4), Inches(2), Inches(8.5), Inches(5))
            tf_body = content_box.text_frame
            tf_body.text = item["content"]
            tf_body.word_wrap = True
            for p in tf_body.paragraphs:
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(80, 80, 80)
                p.space_after = Pt(10)

        # -----------------------------------------------
        # 風格 B: 科技邊框 (Cyberpunk HUD)
        # 特色：黑底、霓虹綠、上下有裝飾線條
        # -----------------------------------------------
        elif style_mode == "cyberpunk":
            # 1. 設定全黑背景
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(10, 10, 15)

            # 2. 畫頂部裝飾條 (霓虹綠)
            bar = shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 
                Inches(0.5), Inches(1.2), Inches(12.33), Inches(0.05)
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(0, 255, 127)
            bar.line.fill.background()

            # 3. 標題
            title_box = shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12), Inches(1))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 255, 127) # 霓虹綠
            p.font.name = "Consolas" # 等寬字體更有科技感

            # 4. 內文
            content_box = shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(5.5))
            tf_body = content_box.text_frame
            tf_body.text = item["content"]
            tf_body.word_wrap = True
            for p in tf_body.paragraphs:
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(220, 220, 220)
                p.font.name = "Consolas"

        # -----------------------------------------------
        # 風格 C: 優雅底線 (Elegant Line)
        # 特色：置中對齊、標題下方有短線、襯線字體
        # -----------------------------------------------
        elif style_mode == "elegant":
            # 背景預設白

            # 1. 標題 (置中)
            title_box = shapes.add_textbox(Inches(1), Inches(1), Inches(11.33), Inches(1))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(44)
            p.font.color.rgb = RGBColor(100, 100, 100) # 質感灰
            p.font.name = "Georgia" # 襯線字體

            # 2. 裝飾短線 (金色，畫在標題下方中央)
            line = shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 
                Inches(6.16), Inches(2.2), Inches(1), Inches(0.05) # 置中計算
            )
            line.fill.solid()
            line.fill.fore_color.rgb = RGBColor(184, 134, 11) # 金色
            line.line.fill.background()

            # 3. 內文 (置中或靠左視內容而定，這裡設靠左但版面縮進)
            content_box = shapes.add_textbox(Inches(2), Inches(3), Inches(9.33), Inches(4))
            tf_body = content_box.text_frame
            tf_body.text = item["content"]
            tf_body.word_wrap = True
            for p in tf_body.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(22)
                p.font.color.rgb = RGBColor(60, 60, 60)
                p.font.name = "Georgia"

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. Streamlit UI
# ==========================================
st.set_page_config(page_title="AI PPT 進階重設計", layout="wide", page_icon="🎨")

st.title("🎨 HW 5-3: AI PPT 進階版型重構")
st.markdown("""
本系統不只替換顏色，更會使用 **Python 幾何繪圖** 重新定義版面結構。
""")
st.divider()

# 側邊欄控制
with st.sidebar:
    st.header("1. 資料來源")
    uploaded_file = st.file_uploader("上傳 PPTX", type="pptx")
    use_sample = st.button("或是：載入範例簡報")
    
    st.divider()
    st.info("💡 提示：Cyberpunk 風格會將字體改為等寬字，適合程式碼展示。")

# 邏輯處理
content_data = None
source_name = ""

if uploaded_file:
    content_data = extract_content_from_pptx(uploaded_file)
    source_name = uploaded_file.name
elif use_sample:
    # 確保你有上傳 sample_presentation.pptx 到 github
    if os.path.exists("sample_presentation.pptx"):
        content_data = extract_content_from_pptx("sample_presentation.pptx")
        source_name = "範例簡報"
    else:
        st.error("找不到 sample_presentation.pptx")

# 顯示結果
if content_data:
    st.subheader(f"✅ 專案來源：{source_name}")
    st.write(f"已提取 {len(content_data)} 頁內容，請選擇版型進行重構：")
    
    # 三欄展示三種風格
    col1, col2, col3 = st.columns(3)

    # Style A
    with col1:
        st.container(border=True)
        st.markdown("### 🏢 現代側邊欄")
        st.caption("雜誌風格排版，左側具備視覺引導色塊。")
        ppt_a = create_styled_pptx(content_data, "sidebar")
        st.download_button("下載 Style A", ppt_a, "style_sidebar.pptx")

    # Style B
    with col2:
        st.container(border=True)
        st.markdown("### 👾 賽博科技 HUD")
        st.caption("深色模式，搭配裝飾性線條與等寬字體。")
        ppt_b = create_styled_pptx(content_data, "cyberpunk")
        st.download_button("下載 Style B", ppt_b, "style_cyberpunk.pptx")

    # Style C
    with col3:
        st.container(border=True)
        st.markdown("### ✒️ 優雅襯線")
        st.caption("大量留白與中央裝飾線，適合高端展示。")
        ppt_c = create_styled_pptx(content_data, "elegant")
        st.download_button("下載 Style C", ppt_c, "style_elegant.pptx")

else:
    if not uploaded_file and not use_sample:
        st.info("👈 請從左側開始")