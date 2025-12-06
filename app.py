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
# 2. 進階樣式生成邏輯
# ==========================================
def create_styled_pptx(data, style_mode):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for item in data:
        slide_layout = prs.slide_layouts[6] 
        slide = prs.slides.add_slide(slide_layout)
        shapes = slide.shapes

        # Style A: 現代側邊欄
        if style_mode == "sidebar":
            sidebar = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(3.5), Inches(7.5))
            sidebar.fill.solid()
            sidebar.fill.fore_color.rgb = RGBColor(44, 62, 80)
            sidebar.line.fill.background()

            title_box = shapes.add_textbox(Inches(4), Inches(0.5), Inches(8.5), Inches(1.5))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = RGBColor(44, 62, 80)
            p.font.name = "Arial Black"

            content_box = shapes.add_textbox(Inches(4), Inches(2), Inches(8.5), Inches(5))
            tf_body = content_box.text_frame
            tf_body.text = item["content"]
            tf_body.word_wrap = True
            for p in tf_body.paragraphs:
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(80, 80, 80)
                p.space_after = Pt(10)

        # Style B: 科技邊框
        elif style_mode == "cyberpunk":
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(10, 10, 15)

            bar = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.2), Inches(12.33), Inches(0.05))
            bar.fill.solid()
            bar.fill.fore_color.rgb = RGBColor(0, 255, 127)
            bar.line.fill.background()

            title_box = shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12), Inches(1))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = RGBColor(0, 255, 127)
            p.font.name = "Consolas"

            content_box = shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(5.5))
            tf_body = content_box.text_frame
            tf_body.text = item["content"]
            tf_body.word_wrap = True
            for p in tf_body.paragraphs:
                p.font.size = Pt(20)
                p.font.color.rgb = RGBColor(220, 220, 220)
                p.font.name = "Consolas"

        # Style C: 優雅底線
        elif style_mode == "elegant":
            title_box = shapes.add_textbox(Inches(1), Inches(1), Inches(11.33), Inches(1))
            tf = title_box.text_frame
            tf.text = item["title"]
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.font.size = Pt(44)
            p.font.color.rgb = RGBColor(100, 100, 100)
            p.font.name = "Georgia"

            line = shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.16), Inches(2.2), Inches(1), Inches(0.05))
            line.fill.solid()
            line.fill.fore_color.rgb = RGBColor(184, 134, 11)
            line.line.fill.background()

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
# 3. Streamlit UI (修復狀態保存問題)
# ==========================================
st.set_page_config(page_title="AI PPT 進階重設計", layout="wide", page_icon="🎨")

st.title("🎨 HW 5-3: AI PPT 進階版型重構")
st.markdown("本系統不只替換顏色，更會使用 **Python 幾何繪圖** 重新定義版面結構。")
st.divider()

# 初始化 Session State (如果沒有的話，先建立空的)
if 'ppt_content' not in st.session_state:
    st.session_state['ppt_content'] = None
if 'source_name' not in st.session_state:
    st.session_state['source_name'] = ""

# 側邊欄控制
with st.sidebar:
    st.header("1. 資料來源")
    
    # 上傳檔案 (注意：上傳檔案會自動更新，因為 file_uploader 本身有狀態)
    uploaded_file = st.file_uploader("上傳 PPTX", type="pptx")
    
    # 載入範例按鈕
    use_sample = st.button("或是：載入範例簡報")
    
    st.divider()
    st.info("💡 提示：Cyberpunk 風格會將字體改為等寬字，適合程式碼展示。")

# 邏輯處理：決定現在要用誰的資料

# 情況 A：使用者剛上傳了檔案 -> 優先使用上傳的檔案
if uploaded_file:
    # 為了避免重複解析，可以檢查是否已經是這個檔案
    if st.session_state['source_name'] != uploaded_file.name:
        st.session_state['ppt_content'] = extract_content_from_pptx(uploaded_file)
        st.session_state['source_name'] = uploaded_file.name

# 情況 B：使用者點擊了「載入範例」 -> 強制切換到範例資料
elif use_sample:
    if os.path.exists("sample_presentation.pptx"):
        st.session_state['ppt_content'] = extract_content_from_pptx("sample_presentation.pptx")
        st.session_state['source_name'] = "範例簡報"
    else:
        st.error("找不到 sample_presentation.pptx")

# 顯示結果 (從 Session State 讀取資料，這樣按下載按鈕時，資料不會不見)
if st.session_state['ppt_content']:
    st.subheader(f"✅ 專案來源：{st.session_state['source_name']}")
    st.write(f"已提取 {len(st.session_state['ppt_content'])} 頁內容，請選擇版型進行重構：")
    
    col1, col2, col3 = st.columns(3)

    # Style A
    with col1:
        st.container(border=True)
        st.markdown("### 🏢 現代側邊欄")
        st.caption("雜誌風格排版，左側具備視覺引導色塊。")
        ppt_a = create_styled_pptx(st.session_state['ppt_content'], "sidebar")
        st.download_button("下載 Style A", ppt_a, "style_sidebar.pptx")

    # Style B
    with col2:
        st.container(border=True)
        st.markdown("### 👾 賽博科技 HUD")
        st.caption("深色模式，搭配裝飾性線條與等寬字體。")
        ppt_b = create_styled_pptx(st.session_state['ppt_content'], "cyberpunk")
        st.download_button("下載 Style B", ppt_b, "style_cyberpunk.pptx")

    # Style C
    with col3:
        st.container(border=True)
        st.markdown("### ✒️ 優雅襯線")
        st.caption("大量留白與中央裝飾線，適合高端展示。")
        ppt_c = create_styled_pptx(st.session_state['ppt_content'], "elegant")
        st.download_button("下載 Style C", ppt_c, "style_elegant.pptx")

else:
    if not uploaded_file:
        st.info("👈 請從左側開始，上傳檔案或載入範例。")