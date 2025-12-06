import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from io import BytesIO

# ==========================================
# 1. 定義五種風格參數 (Rule-Based Style Config)
# ==========================================
STYLES = {
    "minimalist": {
        "name": "極簡白 (Minimalist)",
        "desc": "適合正式商務、學術報告。白底深藍字，強調專業感。",
        "bg_color": RGBColor(255, 255, 255),
        "title_color": RGBColor(0, 51, 102),  # 深海藍
        "text_color": RGBColor(60, 60, 60),   # 深灰
        "font_bold": True
    },
    "cyberpunk": {
        "name": "賽博龐克 (Cyberpunk)",
        "desc": "適合黑客松、科技 Demo。深黑底配霓虹綠，未來感強烈。",
        "bg_color": RGBColor(20, 20, 25),     # 極深灰
        "title_color": RGBColor(0, 255, 150), # 霓虹綠
        "text_color": RGBColor(240, 240, 240),# 亮白
        "font_bold": True
    },
    "warm_paper": {
        "name": "溫暖紙質感 (Warm Paper)",
        "desc": "適合人文、閱讀類主題。米色背景配深褐字，閱讀舒適。",
        "bg_color": RGBColor(250, 245, 230),  # 米黃色 (Linen)
        "title_color": RGBColor(101, 67, 33), # 深褐色
        "text_color": RGBColor(80, 50, 20),   # 咖啡色
        "font_bold": False
    },
    "dark_luxury": {
        "name": "暗夜奢華 (Dark Luxury)",
        "desc": "適合高端產品介紹。黑底配金字，展現高級感。",
        "bg_color": RGBColor(0, 0, 0),        # 純黑
        "title_color": RGBColor(212, 175, 55),# 金色
        "text_color": RGBColor(200, 200, 200),# 淺灰
        "font_bold": True
    },
    "corporate_blue": {
        "name": "穩重藍調 (Corporate Blue)",
        "desc": "適合傳統企業內部報告。全藍背景配白字，穩重不失誤。",
        "bg_color": RGBColor(44, 62, 80),     # 普魯士藍
        "title_color": RGBColor(255, 255, 255), # 白
        "text_color": RGBColor(236, 240, 241),  # 雲朵白
        "font_bold": True
    }
}

# ==========================================
# 2. 核心邏輯函數
# ==========================================

def extract_content_from_pptx(uploaded_file):
    """從上傳的 PPT 提取標題與內文"""
    prs = Presentation(uploaded_file)
    extracted_data = []

    for slide in prs.slides:
        slide_content = {"title": "", "content": ""}
        
        # 抓取標題
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            slide_content["title"] = slide.shapes.title.text_frame.text
        
        # 抓取內文 (搜尋 Placeholders)
        # 策略：尋找除了標題之外，第一個包含文字的框
        for shape in slide.placeholders:
            # idx 1 通常是標準版型的內文框
            if shape.placeholder_format.idx == 1 and shape.has_text_frame:
                slide_content["content"] = shape.text_frame.text
                break # 找到一個就停，避免抓到頁碼或日期
        
        # 只有當該頁有內容時才加入
        if slide_content["title"] or slide_content["content"]:
            extracted_data.append(slide_content)
            
    return extracted_data

def create_styled_pptx(data, style_key):
    """根據指定的 style_key 生成 PPT"""
    prs = Presentation()
    style_config = STYLES[style_key]
    
    for item in data:
        # 使用 Layout 1 (標題 + 內容)
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # --- A. 設定背景顏色 ---
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = style_config["bg_color"]

        # --- B. 設定標題 ---
        title = slide.shapes.title
        title.text = item["title"]
        
        for paragraph in title.text_frame.paragraphs:
            paragraph.font.color.rgb = style_config["title_color"]
            paragraph.font.bold = style_config["font_bold"]
            paragraph.font.name = "Arial" # 通用字體

        # --- C. 設定內文 ---
        if len(slide.placeholders) > 1:
            body = slide.placeholders[1]
            body.text = item["content"]
            
            for paragraph in body.text_frame.paragraphs:
                paragraph.font.color.rgb = style_config["text_color"]
                paragraph.font.size = Pt(20)
                paragraph.font.name = "Arial"

    # 轉為 BytesIO 物件回傳 (不存硬碟)
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
請上傳您的 `.pptx` 檔案進行測試。
""")

st.divider()

# 檔案上傳區
uploaded_file = st.file_uploader("📂 請上傳 PPTX 檔案", type="pptx")

if uploaded_file is not None:
    try:
        # 1. 提取內容
        with st.spinner('正在解析簡報內容...'):
            content_data = extract_content_from_pptx(uploaded_file)
        
        if not content_data:
            st.error("❌ 無法提取內容，請確認您的 PPT 使用了標準的「標題+內文」版型。")
        else:
            st.success(f"✅ 成功提取 {len(content_data)} 頁內容！請選擇下方風格下載：")
            st.divider()

            # 2. 顯示五種風格下載選項
            # 使用 Grid 排版
            cols = st.columns(2) # 兩欄排列
            
            # 遍歷所有風格並生成按鈕
            for index, (key, config) in enumerate(STYLES.items()):
                # 決定放左欄還是右欄
                col = cols[index % 2]
                
                with col:
                    with st.container(border=True):
                        st.subheader(config["name"])
                        st.caption(config["desc"])
                        
                        # 即時生成該風格的 PPT
                        # 注意：這裡雖然是即時生成，但因為 python-pptx 很快，所以不會卡頓
                        ppt_file = create_styled_pptx(content_data, key)
                        
                        st.download_button(
                            label=f"⬇️ 下載 {config['name']} PPT",
                            data=ppt_file,
                            file_name=f"redesigned_{key}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )

    except Exception as e:
        st.error(f"系統發生錯誤：{e}")