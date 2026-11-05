import streamlit as st
from groq import Groq
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE
from io import BytesIO
import json
import re


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="AI Presentation Maker",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

    .stApp {
        background:
            radial-gradient(circle at top left,
                rgba(99, 102, 241, 0.15),
                transparent 35%),
            radial-gradient(circle at bottom right,
                rgba(168, 85, 247, 0.12),
                transparent 35%),
            #0b1020;
    }

    .main-title {
        font-size: 52px;
        font-weight: 800;
        text-align: center;
        margin-top: 20px;
        margin-bottom: 8px;
        color: #ffffff;
    }

    .subtitle {
        text-align: center;
        color: #aeb8d4;
        font-size: 18px;
        margin-bottom: 35px;
    }

    .feature-card {
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 18px;
        padding: 22px;
        height: 100%;
        backdrop-filter: blur(10px);
    }

    .feature-card h3 {
        color: white;
        margin-bottom: 8px;
    }

    .feature-card p {
        color: #b7bfd5;
    }

    .section-title {
        color: white;
        font-size: 24px;
        font-weight: 700;
        margin-top: 25px;
        margin-bottom: 15px;
    }

    .success-box {
        padding: 18px;
        border-radius: 15px;
        background: rgba(34,197,94,0.12);
        border: 1px solid rgba(34,197,94,0.3);
        color: #d8ffe4;
    }

    div.stButton > button {
        border-radius: 12px;
        border: none;
        font-weight: 700;
        min-height: 48px;
    }

    .small-text {
        color: #9ca8c7;
        font-size: 13px;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# COLORS
# =========================================================

COLORS = {
    "dark": RGBColor(15, 23, 42),
    "navy": RGBColor(15, 23, 42),
    "white": RGBColor(255, 255, 255),
    "light": RGBColor(241, 245, 249),
    "gray": RGBColor(100, 116, 139),
    "purple": RGBColor(124, 58, 237),
    "blue": RGBColor(37, 99, 235),
    "cyan": RGBColor(6, 182, 212),
    "pink": RGBColor(219, 39, 119),
    "green": RGBColor(22, 163, 74),
    "orange": RGBColor(234, 88, 12),
}


# =========================================================
# GET GROQ CLIENT
# =========================================================

def get_groq_client():

    try:
        api_key = st.secrets["GROQ_API_KEY"]
    except Exception:
        st.error(
            "GROQ_API_KEY is not configured. "
            "Add it to Streamlit Secrets."
        )
        st.stop()

    return Groq(api_key=api_key)


# =========================================================
# AI GENERATION
# =========================================================

def generate_presentation(
    topic,
    slide_count,
    audience,
    tone,
    language,
    presentation_style
):

    client = get_groq_client()

    prompt = f"""
You are an expert presentation designer and business content writer.

Create a professional PowerPoint presentation.

TOPIC:
{topic}

NUMBER OF SLIDES:
{slide_count}

AUDIENCE:
{audience}

TONE:
{tone}

LANGUAGE:
{language}

DESIGN STYLE:
{presentation_style}

Requirements:

1. Create exactly {slide_count} slides.
2. Give every slide a strong professional title.
3. Keep slide content concise and presentation-friendly.
4. Do not write huge paragraphs.
5. Use bullet points whenever appropriate.
6. Include a logical flow:
   - Introduction
   - Main concepts
   - Important details
   - Examples or applications
   - Conclusion
7. The presentation should feel professionally designed.
8. Content must be useful and informative.
9. Avoid repetitive slides.
10. Use simple language suitable for the selected audience.

Return ONLY valid JSON in this format:

{{
    "presentation_title": "Main presentation title",
    "subtitle": "Short subtitle",
    "slides": [
        {{
            "title": "Slide title",
            "type": "title",
            "content": [
                "Content item 1",
                "Content item 2"
            ],
            "speaker_notes": "Short speaker notes"
        }}
    ]
}}

For the first slide use type "title".

For normal content slides use:
"title", "content"

For a conclusion slide, use:
"type": "conclusion"

Do not include Markdown.
"""

    response = client.chat.completions.create(
       model="openai/gpt-oss-120b",
        messages=[
            {
                "role": "system",
                "content": "You create structured professional presentations."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        temperature=0.7,
        max_tokens=8000
    )

    raw = response.choices[0].message.content.strip()

    # Remove accidental markdown fences
    raw = re.sub(r"```json", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"```", "", raw)

    raw = raw.strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:

        # Try extracting JSON object
        start = raw.find("{")
        end = raw.rfind("}")

        if start == -1 or end == -1:
            raise ValueError("AI returned invalid presentation data.")

        data = json.loads(raw[start:end + 1])

    return data


# =========================================================
# PPT HELPERS
# =========================================================

def add_background(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_top_bar(slide, color):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        Inches(13.333),
        Inches(0.12)
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_footer(slide, slide_number, color=COLORS["gray"]):

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.6),
        Inches(7.05),
        Inches(12.1),
        Inches(0.01)
    )

    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()

    textbox = slide.shapes.add_textbox(
        Inches(11.8),
        Inches(7.1),
        Inches(0.8),
        Inches(0.3)
    )

    tf = textbox.text_frame
    p = tf.paragraphs[0]
    p.text = str(slide_number)
    p.font.size = Pt(10)
    p.font.color.rgb = color
    p.alignment = PP_ALIGN.RIGHT


def add_text(
    slide,
    text,
    left,
    top,
    width,
    height,
    font_size=24,
    color=COLORS["dark"],
    bold=False,
    alignment=PP_ALIGN.LEFT
):

    textbox = slide.shapes.add_textbox(
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height)
    )

    tf = textbox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = alignment

    return textbox


def add_bullets(
    slide,
    items,
    left,
    top,
    width,
    height,
    font_size=20,
    color=COLORS["dark"]
):

    textbox = slide.shapes.add_textbox(
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height)
    )

    tf = textbox.text_frame
    tf.word_wrap = True

    tf.clear()

    for index, item in enumerate(items):

        if index == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.text = str(item)
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(12)
        p.level = 0

        p.text = "• " + str(item)

    return textbox


# =========================================================
# TITLE SLIDE
# =========================================================

def create_title_slide(prs, title, subtitle, style):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_background(slide, COLORS["navy"])

    # Decorative circles
    circle1 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(9.5),
        Inches(-1.5),
        Inches(5),
        Inches(5)
    )

    circle1.fill.solid()
    circle1.fill.fore_color.rgb = COLORS["purple"]
    circle1.fill.transparency = 20
    circle1.line.fill.background()

    circle2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(-2),
        Inches(5),
        Inches(4),
        Inches(4)
    )

    circle2.fill.solid()
    circle2.fill.fore_color.rgb = COLORS["cyan"]
    circle2.fill.transparency = 25
    circle2.line.fill.background()

    # Small label
    add_text(
        slide,
        "AI PRESENTATION MAKER",
        0.8,
        1.2,
        5,
        0.5,
        13,
        COLORS["cyan"],
        True
    )

    # Main title
    add_text(
        slide,
        title,
        0.8,
        2.0,
        9.5,
        1.7,
        38,
        COLORS["white"],
        True
    )

    # Subtitle
    add_text(
        slide,
        subtitle,
        0.8,
        4.0,
        8.5,
        1.0,
        20,
        RGBColor(203, 213, 225),
        False
    )

    # Style badge
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.8),
        Inches(5.4),
        Inches(2.3),
        Inches(0.55)
    )

    badge.fill.solid()
    badge.fill.fore_color.rgb = COLORS["purple"]
    badge.line.fill.background()

    add_text(
        slide,
        style.upper(),
        0.95,
        5.48,
        2,
        0.35,
        10,
        COLORS["white"],
        True,
        PP_ALIGN.CENTER
    )

    return slide


# =========================================================
# CONTENT SLIDE
# =========================================================

def create_content_slide(prs, slide_data, slide_number):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_background(slide, COLORS["white"])
    add_top_bar(slide, COLORS["purple"])

    # Slide title
    add_text(
        slide,
        slide_data.get("title", "Untitled"),
        0.7,
        0.55,
        11.5,
        0.7,
        28,
        COLORS["navy"],
        True
    )

    # Accent line
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.72),
        Inches(1.35),
        Inches(1.0),
        Inches(0.07)
    )

    accent.fill.solid()
    accent.fill.fore_color.rgb = COLORS["purple"]
    accent.line.fill.background()

    # Content card
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.7),
        Inches(1.75),
        Inches(11.9),
        Inches(4.7)
    )

    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(248, 250, 252)
    card.line.color.rgb = RGBColor(226, 232, 240)

    content = slide_data.get("content", [])

    add_bullets(
        slide,
        content,
        1.1,
        2.1,
        10.9,
        3.9,
        21,
        COLORS["dark"]
    )

    add_footer(slide, slide_number)

    return slide


# =========================================================
# CONCLUSION SLIDE
# =========================================================

def create_conclusion_slide(prs, slide_data, slide_number):

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    add_background(slide, COLORS["navy"])

    # Main decorative shape
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(8.8),
        Inches(-1.5),
        Inches(5.5),
        Inches(5.5)
    )

    circle.fill.solid()
    circle.fill.fore_color.rgb = COLORS["purple"]
    circle.fill.transparency = 25
    circle.line.fill.background()

    add_text(
        slide,
        "KEY TAKEAWAYS",
        0.8,
        1.1,
        4,
        0.5,
        14,
        COLORS["cyan"],
        True
    )

    add_text(
        slide,
        slide_data.get("title", "Conclusion"),
        0.8,
        1.8,
        8.5,
        1.0,
        34,
        COLORS["white"],
        True
    )

    content = slide_data.get("content", [])

    add_bullets(
        slide,
        content,
        0.9,
        3.0,
        8.5,
        2.5,
        20,
        RGBColor(226, 232, 240)
    )

    add_footer(
        slide,
        slide_number,
        RGBColor(148, 163, 184)
    )

    return slide


# =========================================================
# CREATE POWERPOINT
# =========================================================

def create_powerpoint(data, style):

    prs = Presentation()

    # Widescreen 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    title = data.get(
        "presentation_title",
        "AI Generated Presentation"
    )

    subtitle = data.get(
        "subtitle",
        "Created with AI Presentation Maker"
    )

    create_title_slide(
        prs,
        title,
        subtitle,
        style
    )

    slides = data.get("slides", [])

    # Avoid duplicate title slide
    content_slides = slides[1:] if len(slides) > 1 else slides

    slide_number = 2

    for slide_data in content_slides:

        slide_type = slide_data.get("type", "content")

        if slide_type == "conclusion":
            create_conclusion_slide(
                prs,
                slide_data,
                slide_number
            )
        else:
            create_content_slide(
                prs,
                slide_data,
                slide_number
            )

        slide_number += 1

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown("## ⚙️ Presentation Settings")

    audience = st.selectbox(
        "Audience",
        [
            "Students",
            "Teachers",
            "Business Professionals",
            "General Audience",
            "Researchers",
            "Clients"
        ]
    )

    tone = st.selectbox(
        "Tone",
        [
            "Professional",
            "Educational",
            "Creative",
            "Corporate",
            "Simple",
            "Persuasive"
        ]
    )

    language = st.selectbox(
        "Language",
        [
            "English",
            "Urdu",
            "Roman Urdu",
            "Arabic",
            "Spanish"
        ]
    )

    presentation_style = st.selectbox(
        "Design Style",
        [
            "Modern",
            "Corporate",
            "Minimal",
            "Creative",
            "Elegant",
            "Technology"
        ]
    )

    slide_count = st.slider(
        "Number of Slides",
        min_value=3,
        max_value=15,
        value=8
    )

    st.markdown("---")

    st.markdown(
        """
        <div class="small-text">
        🔐 Your Groq API key is stored securely
        using Streamlit Secrets.
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# MAIN UI
# =========================================================

st.markdown(
    '<div class="main-title">✨ AI Presentation Maker</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Create professional PowerPoint presentations with AI in seconds.'
    '</div>',
    unsafe_allow_html=True
)


# Feature cards

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(
        """
        <div class="feature-card">
            <h3>🤖 AI Powered</h3>
            <p>
            Generate structured presentation content
            using advanced AI.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

with col2:
    st.markdown(
        """
        <div class="feature-card">
            <h3>🎨 Professional Design</h3>
            <p>
            Generate clean, modern and presentation-ready
            PowerPoint slides.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

with col3:
    st.markdown(
        """
        <div class="feature-card">
            <h3>📥 Download PPTX</h3>
            <p>
            Download your generated presentation
            as a PowerPoint file.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )


st.markdown(
    '<div class="section-title">📝 Create Your Presentation</div>',
    unsafe_allow_html=True
)


topic = st.text_area(
    "Presentation Topic",
    placeholder=(
        "Example: Artificial Intelligence in Education"
    ),
    height=120
)


if st.button(
    "🚀 Generate Professional Presentation",
    use_container_width=True
):

    if not topic.strip():
        st.warning("Please enter a presentation topic.")
        st.stop()

    with st.spinner(
        "🤖 AI is creating your professional presentation..."
    ):

        try:

            data = generate_presentation(
                topic=topic,
                slide_count=slide_count,
                audience=audience,
                tone=tone,
                language=language,
                presentation_style=presentation_style
            )

            ppt_file = create_powerpoint(
                data,
                presentation_style
            )

            st.session_state["ppt_file"] = ppt_file.getvalue()
            st.session_state["presentation_data"] = data

            st.success(
                "🎉 Presentation generated successfully!"
            )

        except Exception as e:

            st.error(
                f"Presentation generation failed: {str(e)}"
            )


# =========================================================
# RESULT
# =========================================================

if "presentation_data" in st.session_state:

    data = st.session_state["presentation_data"]

    st.markdown("---")

    st.markdown(
        '<div class="section-title">📊 Presentation Preview</div>',
        unsafe_allow_html=True
    )

    st.write(
        f"### {data.get('presentation_title', 'Presentation')}"
    )

    st.write(
        data.get(
            "subtitle",
            "AI Generated Presentation"
        )
    )

    slides = data.get("slides", [])

    for i, slide in enumerate(slides, start=1):

        with st.expander(
            f"Slide {i}: {slide.get('title', 'Untitled')}"
        ):

            content = slide.get("content", [])

            for item in content:
                st.write(f"• {item}")

            if slide.get("speaker_notes"):
                st.caption(
                    f"Speaker Notes: {slide['speaker_notes']}"
                )

    st.markdown("---")

    st.download_button(
        label="📥 Download PowerPoint Presentation",
        data=st.session_state["ppt_file"],
        file_name="AI_Generated_Presentation.pptx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "presentationml.presentation"
        ),
        use_container_width=True
    )
