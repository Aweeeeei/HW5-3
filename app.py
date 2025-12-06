import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from io import BytesIO
import os

# ==========================================
# 1. 定義五種風格參數
# ==========================================
STYLES = {
    "minimalist": {
        "name": "極簡白 (Minimalist)",
        "desc": "適合正式商務、學術報告。白底深藍字，強調專業感。",
        "bg_color": RGBColor(255, 255, 255),
        "title_color": RGBColor(0, 51, 102),
        "text_color": RGBColor(60, 60, 60),
        "font_bold": True
    },
    "cyberpunk": {
        "name": "賽博龐克 (Cyberpunk)",
        "desc": "適合黑客松、科技 Demo。深黑底配霓虹綠，未來感強烈。",
        "bg_color": RGBColor(20, 20, 25),
        "title_color": RGBColor(0, 255, 150),
        "text_color": RGBColor(240, 240, 240),
        "font_bold": True
    },
    "warm_paper": {
        "name": "溫暖紙質感 (Warm Paper)",
        "desc": "適合人文、閱讀類主題。米色背景配深褐字，閱讀舒適。",
        "bg_color": RGBColor(250, 245, 230),
        "title_color": RGBColor(101, 67, 33),
        "text_color": RGBColor(80, 50, 20),
        "font_bold": False
    },
    "dark_luxury": {
        "name": "暗夜奢華 (Dark Luxury)",
        "desc": "適合高端產品介紹。黑底配金字，展現高級感。",
        "bg_color": RGBColor(0, 0, 0),
        "title_color": RGBColor(212, 175, 55),
        "text_color": RGBColor(200, 200, 200),
        "font_bold": True
    },
    "corporate_blue": {
        "name": "穩重藍調 (Corporate Blue)",
        "desc": "適合傳統企業內部報告。全藍背景配白字，穩重不失誤。",
        "bg_color": RGBColor(44, 62, 80),
        "title_color": RGBColor(255, 255, 255),
        "text_color": RGBColor(236, 240, 241),
        "font_bold": True
    }
}

# ==========================================
# 2. 核心邏輯函數
# ==========================================

def extract_content_from_pptx(file_obj):
    """從上傳的檔案物件或路徑提取內容"""
    prs = Presentation(file_obj)
    extracted_data = []

    for slide in prs.slides:
        slide_content = {"title": "", "content": ""}
        
        # 抓取標題
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            slide_content["title"] = slide.shapes.title.text_frame.text
        
        # 抓取內文
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1 and shape.has_text_frame:
                slide_content["content"] = shape.text_frame.text
                break 
        
        if slide_content["title"] or slide_content["content"]:
            extracted_data.append(slide_content)
            
    return extracted_data

def create_styled_pptx(data, style_key):
    """根據指定的 style_key 生成 PPT"""
    prs = Presentation()
    style_config = STYLES[style_key]
    
    for item in data:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # 設定背景
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = style_config["bg_color"]

        # 設定標題
        title = slide.shapes.title
        title.text = item["title"]
        for paragraph in title.text_frame.paragraphs:
            paragraph.font.color.rgb = style_config["title_color"]
            paragraph.font.bold = style_config["font_bold"]
            paragraph.font.name = "Arial"

        # 設定內文
        if len(slide.placeholders) > 1:
            body = slide.placeholders[1]
            body.text = item["content"]
            for paragraph in body.text_frame.paragraphs:
                paragraph.font.color.rgb = style_config["text_color"]
                paragraph.font.size = Pt(20)
                paragraph.font.name = "Arial"

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. Streamlit 介面
# ==========================================
st.set_page_config(page_title="AI PPT 風格重塑", layout="wide", page_icon="📊")

st.title("📊 HW 5-3: AI PPT 風格自動重塑系統")
st.markdown("""
本工具使用 **Python-pptx** 自動化技術，將您的簡報內容提取後，重新注入 **5 種不同的設計風格**。
""")

st.divider()

# --- 側邊欄或主要區域：選擇來源 ---
col1, col2 = st.columns([1, 1], gap="large")

content_data = None
current_source_name = ""

with col1:
    st.subheader("方式 A：上傳您的檔案")
    uploaded_file = st.file_uploader("請上傳 PPTX 檔案", type="pptx")

with col2:
    st.subheader("方式 B：使用測試範例")
    st.write("手邊沒有簡報？直接載入範例試試看！")
    use_sample = st.button("📂 載入範例簡報 (sample.pptx)", use_container_width=True)

# --- 處理邏輯 ---
try:
    if uploaded_file is not None:
        content_data = extract_content_from_pptx(uploaded_file)
        current_source_name = uploaded_file.name
        
    elif use_sample:
        sample_path = "sample_presentation.pptx"
        if os.path.exists(sample_path):
            content_data = extract_content_from_pptx(sample_path)
            current_source_name = "sample_presentation.pptx"
            st.session_state['use_sample_active'] = True # 保持狀態
        else:
            st.error("⚠️ 找不到範例檔案，請確認 sample_presentation.pptx 是否在 Github 倉庫中。")
    
    # 為了讓「使用範例」在點擊下載按鈕後不消失，可以使用 session_state (選用，簡單版可略過)
    # 如果使用者剛剛點過範例，且沒有上傳新檔案，就維持範例資料
    if content_data is None and st.session_state.get('use_sample_active') and uploaded_file is None:
        sample_path = "sample_presentation.pptx"
        if os.path.exists(sample_path):
            content_data = extract_content_from_pptx(sample_path)
            current_source_name = "sample_presentation.pptx"

    # --- 顯示結果區 ---
    if content_data:
        st.divider()
        st.success(f"✅ 已成功載入：**{current_source_name}** (共 {len(content_data)} 頁)")
        st.caption("請從下方選擇喜歡的風格下載：")

        # 顯示五種風格下載選項 (Grid 排版)
        grid_cols = st.columns(2) 
        
        for index, (key, config) in enumerate(STYLES.items()):
            col = grid_cols[index % 2]
            with col:
                with st.container(border=True):
                    st.subheader(config["name"])
                    st.caption(config["desc"])
                    
                    # 生成 PPT
                    ppt_file = create_styled_pptx(content_data, key)
                    
                    st.download_button(
                        label=f"⬇️ 下載 {config['name']}",
                        data=ppt_file,
                        file_name=f"redesigned_{key}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
                    
except Exception as e:
    st.error(f"系統發生錯誤：{e}")