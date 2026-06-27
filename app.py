"""
ACTA — AI-Powered Research Assistant
============================================
A Streamlit web application for scientific question synthesis
and academic paper framework design.

Dependencies: streamlit, openai, python-docx
"""

import streamlit as st
import os
import re
from openai import OpenAI
from docx import Document

# ==================== Page Configuration ====================
st.set_page_config(
    page_title="ACTA · Research Assistant",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==================== Custom CSS ====================
st.markdown(
    """
<style>
    /* --- Global: minimal slate-blue palette --- */
    .stApp {
        background-color: #f8fafc;
    }

    .main .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 960px !important;
    }

    /* --- Sidebar --- */
    section[data-testid="stSidebar"] {
        background-color: #f1f5f9;
        border-right: 1px solid #e2e8f0;
    }

    section[data-testid="stSidebar"] .block-container {
        padding: 1.5rem 1.2rem;
    }

    /* Sidebar radio buttons */
    div[role="radiogroup"] label {
        padding: 10px 14px !important;
        border-radius: 6px !important;
        margin-bottom: 6px !important;
        font-size: 15px !important;
        font-weight: 500;
        transition: background 0.15s;
    }

    div[role="radiogroup"] label:hover {
        background: #e2e8f0 !important;
    }

    /* Hide radio indicators */
    div[role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    div[role="radiogroup"] label {
        padding-left: 14px !important;
    }

    /* --- Hero-style page titles --- */
    .page-hero {
        font-size: 36px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: -0.02em;
        margin-bottom: 4px;
        line-height: 1.2;
    }

    .page-tagline {
        font-size: 16px;
        color: #64748b;
        font-weight: 400;
        margin-bottom: 32px;
    }

    /* --- Input fields --- */
    .stTextInput > div > div > input,
    .stTextArea textarea {
        border-radius: 6px !important;
        border: 1px solid #cbd5e1 !important;
        padding: 10px 14px !important;
        font-size: 15px !important;
        background: #ffffff !important;
    }

    .stTextInput > div > div > input:focus,
    .stTextArea textarea:focus {
        border-color: #3b82f6 !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.12) !important;
    }

    .stSelectbox > div > div {
        border-radius: 6px !important;
        border: 1px solid #cbd5e1 !important;
    }

    /* --- Buttons --- */
    .stButton > button {
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 12px 28px !important;
        border: none !important;
        background: #3b82f6 !important;
        color: #ffffff !important;
        font-size: 15px !important;
        transition: background 0.15s, box-shadow 0.15s !important;
        letter-spacing: 0.01em;
    }

    .stButton > button:hover {
        background: #2563eb !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.28) !important;
    }

    .stButton > button:active {
        background: #1d4ed8 !important;
    }

    /* --- Spinner --- */
    .stSpinner > div {
        border-top-color: #3b82f6 !important;
    }

    /* --- Alerts --- */
    div[data-testid="stAlert"] {
        border-radius: 6px;
        padding: 14px 18px;
        font-size: 14px;
    }

    /* --- Dividers --- */
    hr {
        border: none;
        border-top: 1px solid #e2e8f0;
        margin: 32px 0;
    }

    /* --- Question blocks (page 1) --- */
    .question-block {
        background: #f8fafc;
        border-left: 4px solid #3b82f6;
        border-radius: 0 6px 6px 0;
        padding: 18px 22px;
        margin: 16px 0;
    }

    .question-block .q-title {
        font-size: 17px;
        font-weight: 600;
        color: #1e293b;
        line-height: 1.5;
        margin-bottom: 8px;
    }

    .question-block .q-keywords {
        font-size: 13px;
        color: #64748b;
        line-height: 1.6;
    }

    .kw-tag {
        display: inline-block;
        background: #dbeafe;
        color: #1e40af;
        font-size: 12px;
        font-weight: 500;
        border-radius: 4px;
        padding: 2px 8px;
        margin: 2px 4px 2px 0;
    }

    .kw-tag.extended {
        background: #fef3c7;
        color: #92400e;
    }

    /* --- Success banner --- */
    .stSuccess {
        background: #f0fdf4;
        border-left: 4px solid #22c55e;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ==================== Constants ====================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROMPT_FILE_PAGE1 = os.path.join(APP_DIR, "prompts_page1.docx")
PROMPT_FILE_PAGE2 = os.path.join(APP_DIR, "prompts_page2.docx")
PROMPT_FILE_PAGE3 = os.path.join(APP_DIR, "prompts_page3.docx")
PROMPT_FILE_PAGE4 = os.path.join(APP_DIR, "prompts_page4.docx")

# Default prompt templates (auto-created as .docx on first run)
DEFAULT_PROMPT_PAGE1 = (
    "Based on the keywords [{keywords}], generate 5 forward-looking "
    "scientific research questions. "
    "For each question, provide:\n"
    "1. The question itself (beginning with 'How', 'In what ways', "
    "or 'To what extent')\n"
    "2. A set of extended keywords derived from the user's input, "
    "broadening the research scope\n\n"
    "Format each question as:\n"
    "---QUESTION---\n"
    "Q: [question text]\n"
    "EXTENDED: [comma-separated extended keywords]\n"
    "---END---"
)

DEFAULT_PROMPT_PAGE2 = (
    "Generate a rigorous academic paper outline for the paper titled "
    "\"{title}\".\n"
    "Paper type: [{paper_type}]\n"
    "Additional description: {description}\n\n"
    "Requirements:\n"
    "- The outline must contain at most 6 major chapters (level-1 headings).\n"
    "- Each chapter contains level-2 sections.\n"
    "- Must include: Introduction, Related Work, "
    "Methodology/Experimental Design, Results & Analysis, Discussion, Conclusion.\n"
    "- Use Markdown format: # for chapters, ## for sections, "
    "### for subsections.\n"
    "- Keep headings concise and academically rigorous."
)

SYSTEM_PROMPT = (
    "You are a senior research scientist and academic writing expert. "
    "Provide rigorous, insightful, and well-structured academic output. "
    "Always follow the requested format precisely."
)

DEFAULT_PROMPT_PAGE3 = (
    "You are helping a researcher draft the Introduction section of an academic paper. "
    "Based on the topic description below, generate key argument sentences "
    "(thesis statements) for each of the four standard Introduction paragraphs.\n\n"
    "Topic: {topic}\n\n"
    "Structure your output into exactly four labeled sections:\n\n"
    "**Paragraph 1 — Broad Background & Scientific Problem:**\n"
    "Provide 2–3 key thesis sentences that establish the research domain, "
    "its importance, and the core scientific gap. "
    "Where a claim would normally be backed by a reference, insert a placeholder "
    "in exactly this format: [可插入XX类型文献以强化“xxx”论证] "
    "where 'XX类型文献' describes the reference type needed "
    "(e.g., 综述/原创研究/方法论文/基准数据集) "
    "and 'xxx' is a 3–8 word summary of the specific claim to support.\n\n"
    "**Paragraph 2 — Existing Research:**\n"
    "Provide 3–4 key thesis sentences summarizing the landscape of existing work. "
    "Use citation placeholders as described above at every point where "
    "a reference is expected.\n\n"
    "**Paragraph 3 — Limitations of Existing Techniques:**\n"
    "Provide 2–3 key thesis sentences critically analyzing gaps and shortcomings "
    "in prior work. Use citation placeholders where specific limitations "
    "are attributed to existing studies.\n\n"
    "**Paragraph 4 — Our Contribution:**\n"
    "Provide 2–3 key thesis sentences stating the novel approach and "
    "contributions of this paper. Use placeholders only where absolutely necessary "
    "(e.g., referencing baseline methods for comparison).\n\n"
    "CRITICAL RULES:\n"
    "- Output ONLY thesis/argument sentences — do NOT write full paragraphs.\n"
    "- Every citation-requiring claim MUST use the placeholder format exactly:\n"
    "  [可插入XX类型文献以强化“xxx”论证]\n"
    "- Fill 'XX类型文献' and 'xxx' with context-appropriate content.\n"
    "- Do NOT invent any real or fake paper titles, author names, or DOIs.\n"
    "- Do NOT perform any literature search.\n"
    "- Use formal academic English throughout.\n"
    "- Output only the content — no meta-commentary."
)


DEFAULT_PROMPT_PAGE4 = (
    "# About ACTA\n\n"
    "ACTA (ACademic Trash fActory) is an AI-powered research writing assistant "
    "designed to help researchers draft and refine academic papers.\n\n"
    "## Features\n\n"
    "- **Scientific Question Synthesis** — Generate forward-looking research questions from keywords\n"
    "- **Framework Architect** — Build structured paper outlines with hierarchical chapters\n"
    "- **Introduction Generator** — Draft key thesis sentences for Introduction sections with citation placeholders\n\n"
    "## How to Use\n\n"
    "1. Configure your API key and model in the sidebar\n"
    "2. Navigate to the desired tool page\n"
    "3. Enter your research topic or keywords\n"
    "4. Click **Generate** to get AI-assisted output\n\n"
    "## Prompt Customization\n\n"
    "Each tool page is backed by a Word document (`.docx`) in the project folder. "
    "Edit these files to customize the prompts and fine-tune the output format.\n\n"
    "## Contact\n\n"
    "For questions or feedback, please reach out to the project maintainer."
)




# ==================== Helper Functions ====================

def ensure_prompt_doc_exists(filepath, default_content):
    """Create the prompt Word document if it does not already exist."""
    if not os.path.exists(filepath):
        try:
            doc = Document()
            for para_text in default_content.split("\n"):
                doc.add_paragraph(para_text)
            doc.save(filepath)
        except Exception as e:
            st.warning(
                f"Could not auto-create prompt document "
                f"{os.path.basename(filepath)}: {e}"
            )


def load_prompt_from_docx(filepath, fallback):
    """Load the prompt template from a Word document.

    If the document is missing or empty, it is created from the fallback
    string and then re-read.
    """
    ensure_prompt_doc_exists(filepath, fallback)
    try:
        doc = Document(filepath)
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        if not paragraphs:
            ensure_prompt_doc_exists(filepath, fallback)
            doc = Document(filepath)
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        st.error(f"Failed to read prompt document: {e}")
        return fallback


def call_llm(prompt, system_prompt=SYSTEM_PROMPT):
    """Invoke the LLM API with a 120-second timeout.

    Reads API credentials from st.session_state. Raises ValueError when
    required configuration is missing.
    """
    api_key = st.session_state.get("api_key", "").strip()
    base_url = st.session_state.get("base_url", "").strip()
    model = st.session_state.get("model", "gpt-3.5-turbo").strip()

    if not api_key:
        raise ValueError("Please configure your API Key in the sidebar.")
    if not base_url:
        raise ValueError("Please configure the API Base URL in the sidebar.")
    if not model:
        raise ValueError("Please configure the Model Name in the sidebar.")

    # 创建 HTTP 客户端，设置超时
    import httpx
    http_client = httpx.Client(
        timeout=httpx.Timeout(120.0, connect=60.0),
    )

    client = OpenAI(
        api_key=api_key,
        base_url=base_url,
        http_client=http_client,
    )

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt},
        ],
        temperature=0.7,
        max_tokens=4096,
        # timeout 在这里设置，而不是在 OpenAI() 中
        timeout=httpx.Timeout(120.0, connect=60.0),
    )

    return response.choices[0].message.content

def parse_questions(raw_text):
    """Parse LLM output into structured question entries.

    Expected format per question:
        ---QUESTION---
        Q: <question text>
        KEYWORDS: <comma-separated>
        EXTENDED: <comma-separated>
        ---END---

    Returns a list of dicts with keys: question, keywords, extended.
    """
    entries = []
    blocks = re.split(r"---QUESTION---", raw_text)
    for block in blocks:
        block = block.strip()
        if not block:
            continue

        q_match = re.search(r"Q:\s*(.+?)(?:\n|---END)", block, re.DOTALL)
        kw_match = re.search(
            r"KEYWORDS:\s*(.+?)(?:\n|---END|EXTENDED)",
            block, re.IGNORECASE | re.DOTALL,
        )
        ex_match = re.search(
            r"EXTENDED:\s*(.+?)(?:\n|---END|$)",
            block, re.IGNORECASE | re.DOTALL,
        )

        if q_match:
            question = q_match.group(1).strip()
            keywords_raw = kw_match.group(1).strip() if kw_match else ""
            extended_raw = ex_match.group(1).strip() if ex_match else ""

            keywords = [
                k.strip() for k in keywords_raw.split(",") if k.strip()
            ]
            extended = [
                k.strip() for k in extended_raw.split(",") if k.strip()
            ]

            entries.append({
                "question": question,
                "keywords": keywords,
                "extended": extended,
            })

    return entries


def render_framework_markdown(md_text):
    """Render paper framework with indentation and font-size by heading level.

    Heading mapping:
        #   (level 1) — 26px, bold, no indent   (chapter)
        ##  (level 2) — 21px, bold, 20px indent (section)
        ### (level 3) — 17px, semi-bold, 40px indent (subsection)
        ####+       — 14–15px, 56px indent

    Body paragraphs are rendered with comfortable line-height.
    """
    lines = md_text.strip().split("\n")
    html_parts = []

    heading_styles = {
        1: {"size": 26, "weight": 700, "indent": 0,  "color": "#0f172a"},
        2: {"size": 21, "weight": 700, "indent": 20, "color": "#1e293b"},
        3: {"size": 17, "weight": 600, "indent": 40, "color": "#334155"},
        4: {"size": 15, "weight": 600, "indent": 56, "color": "#475569"},
        5: {"size": 14, "weight": 500, "indent": 56, "color": "#64748b"},
        6: {"size": 13, "weight": 500, "indent": 56, "color": "#64748b"},
    }

    for line in lines:
        stripped = line.strip()
        if not stripped:
            html_parts.append('<div style="height:6px;"></div>')
            continue

        # Detect Markdown heading
        heading_match = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading_match:
            level = len(heading_match.group(1))
            title = heading_match.group(2)
            style = heading_styles.get(level, heading_styles[6])

            html_parts.append(
                f'<div style="'
                f'margin-left:{style["indent"]}px; '
                f'font-size:{style["size"]}px; '
                f'font-weight:{style["weight"]}; '
                f'color:{style["color"]}; '
                f'margin-top:{28 if level <= 2 else 20}px; '
                f'margin-bottom:{10 if level <= 2 else 8}px; '
                f'line-height:1.35; '
                f'letter-spacing:-0.01em;'
                f'">{title}</div>'
            )
            continue

        # Bold standalone paragraph
        bold_match = re.match(r"^\*\*(.+)\*\*$", stripped)
        if bold_match:
            html_parts.append(
                f'<p style="margin-left:20px; font-weight:600; '
                f'color:#334155; line-height:1.7; '
                f'margin-bottom:4px;">{bold_match.group(1)}</p>'
            )
            continue

        # Regular body paragraph
        html_parts.append(
            f'<p style="margin-left:18px; color:#475569; '
            f'line-height:1.8; margin-bottom:6px; '
            f'font-size:15px;">{stripped}</p>'
        )

    return "\n".join(html_parts)


# ==================== Sidebar ====================
with st.sidebar:
    # Hero-style branding — prominent site title
    st.markdown(
        '<div style="font-size:50px; font-weight:800; color:#0f172a; '
        'letter-spacing:-0.02em; margin-bottom:0px;">ACTA</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="font-size:13px; color:#94a3b8; font-style:italic; '
        'letter-spacing:0.05em; margin-bottom:2px;">ACademic Trash fActory</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="font-size:13px; color:#64748b; '
        'margin-bottom:8px;">AI-Powered Research Assistant</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # Page navigation
    page = st.radio(
        "Navigation",
        [
            "🔬 Scientific Question Synthesis",
            "📐 Framework Architect",
            "📝 Introduction Generator",
            "ℹ️ About ACTA",
        ],
        label_visibility="collapsed",
    )

    st.markdown("---")

    # API configuration (collapsible)
    with st.expander("⚙️ API Configuration", expanded=False):
        api_key = st.text_input(
            "API Key",
            type="password",
            value=st.session_state.get("api_key", ""),
            placeholder="sk-...",
            help="Your API key",
        )
        base_url = st.text_input(
            "Base URL",
            value=st.session_state.get(
                "base_url", "https://api.siliconflow.cn/v1"
            ),
            placeholder="https://api.siliconflow.cn/v1",
            help="API endpoint (OpenAI, DeepSeek, local deployment, etc.)",
        )
        model = st.text_input(
            "Model Name",
            value=st.session_state.get("model", "deepseek-ai/DeepSeek-V4-Pro"),
            placeholder="deepseek-ai/DeepSeek-V4-Pro",
            help="e.g. deepseek-ai/DeepSeek-V4-Pro",
        )

        st.session_state["api_key"] = api_key
        st.session_state["base_url"] = base_url
        st.session_state["model"] = model

    st.markdown("---")
    st.caption("ACTA v1.0  ·  Research Assistant")


# ==================== Page 1: Scientific Question Synthesis ====================
if page == "🔬 Scientific Question Synthesis":
    # Prevent the Streamlit toolbar from overlapping the hero title
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="page-hero">🔬 Scientific Question Synthesis</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-tagline">'
        "Enter research keywords (separated by commas or semicolons) and let AI generate forward-looking "
        "scientific questions with associated and extended keyword sets."
        "</div>",
        unsafe_allow_html=True,
    )

    # Input area
    keywords = st.text_input(
        "Research Keywords",
        placeholder="e.g. quantum computing, error correction",
        help="Enter 1–5 keywords, separated by commas",
        key="keywords_input",
    )

    col_left, col_right = st.columns([1, 2])
    with col_left:
        generate_btn = st.button(
            "Generate",
            type="primary",
            use_container_width=True,
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    if generate_btn:
        if not keywords.strip():
            st.warning("⚠️ Please enter at least one research keyword.")
        else:
            try:
                with st.spinner("ACTA is thinking ..."):
                    prompt_template = load_prompt_from_docx(
                        PROMPT_FILE_PAGE1, DEFAULT_PROMPT_PAGE1
                    )
                    final_prompt = prompt_template.replace(
                        "{keywords}", keywords.strip()
                    )

                    raw_result = call_llm(final_prompt)
                    questions = parse_questions(raw_result)
                    st.session_state["page1_questions"] = questions
                    st.session_state["page1_keywords"] = keywords.strip()

                if not questions:
                    st.warning(
                        "⚠️ The model did not return questions in the "
                        "expected format. Raw response shown below."
                    )
                    st.text(raw_result)
                else:
                    st.markdown(
                        "### 📝 Synthesized Questions "
                        f"({len(questions)} results)"
                    )

                    for i, q in enumerate(questions, 1):
                        ex_tags = "".join(
                            f'<span class="kw-tag extended">{k}</span>'
                            for k in q["extended"]
                        )

                        st.markdown(
                            f'<div class="question-block">'
                            f'<div class="q-title">'
                            f"<strong>Q{i}:</strong> {q['question']}"
                            f"</div>"
                            f'<div class="q-keywords">'
                            f"<strong>Extended Keywords:</strong> {ex_tags}"
                            f"</div>"
                            f"</div>",
                            unsafe_allow_html=True,
                        )

                    st.success(
                        f"✅ Successfully synthesized "
                        f"{len(questions)} research questions!"
                    )

            except ValueError as e:
                st.error(f"❌ Configuration Error: {e}")
            except Exception as e:
                st.error(
                    f"❌ API Call Failed: {e}\n\n"
                    "Troubleshooting:\n"
                    "- Verify your API Key is correct\n"
                    "- Check that the Base URL is reachable\n"
                    "- Confirm the model name exists\n"
                    "- Ensure network connectivity is stable"
                )

    # Show cached results when returning to this page
    if "page1_questions" in st.session_state and not generate_btn:
        questions = st.session_state["page1_questions"]
        if questions:
            st.markdown("### Synthesized Questions (cached) " + str(len(questions)) + " results")
            for i, q in enumerate(questions, 1):
                ex_tags = "".join("<span class=kw-tag extended>" + k + "</span>" for k in q["extended"])
                st.markdown(
                    "<div class=question-block>" +
                    "<div class=q-title><strong>Q" + str(i) + ":</strong> " + q["question"] + "</div>" +
                    "<div class=q-keywords><strong>Extended Keywords:</strong> " + ex_tags + "</div>" +
                    "</div>",
                    unsafe_allow_html=True,
                )



# ==================== Page 3: Introduction Generator ====================
elif page == "📝 Introduction Generator":
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="page-hero">📝 Introduction Generator</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-tagline">'
        "Describe your research topic and AI will generate a well-structured "
        "academic Introduction section with background, literature review, "
        "limitations, and contributions."
        "</div>",
        unsafe_allow_html=True,
    )

    # Input area
    topic_description = st.text_area(
        "Research Topic Description",
        placeholder=(
            "Describe your paper’s research topic, key ideas, methodology, "
            "and contributions in detail. The more context you provide, "
            "the better the generated Introduction will be.\n\n"
            "e.g. This paper proposes a novel deep learning framework for "
            "real-time traffic flow prediction using graph neural networks "
            "and attention mechanisms..."
        ),
        help="Provide a comprehensive description of your research topic (max ~500 words)",
        max_chars=4000,
        height=200,
        key="topic_description_input",
    )

    col_left, col_right = st.columns([1, 2])
    with col_left:
        generate_btn = st.button(
            "Generate",
            type="primary",
            use_container_width=True,
            key="generate_intro_btn",
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    if generate_btn:
        if not topic_description.strip():
            st.warning("⚠️ Please enter a research topic description.")
        else:
            try:
                with st.spinner("ACTA is writing the Introduction ..."):
                    prompt_template = load_prompt_from_docx(
                        PROMPT_FILE_PAGE3, DEFAULT_PROMPT_PAGE3
                    )
                    final_prompt = prompt_template.replace(
                        "{topic}", topic_description.strip()
                    )

                    raw_result = call_llm(final_prompt)
                    st.session_state["page3_intro"] = raw_result


                    st.markdown("### 📄 Generated Introduction")
                    # Render the introduction as Markdown
                    st.markdown(raw_result)
                    st.success("✅ Introduction generated successfully!")

            except ValueError as e:
                st.error(f"❌ Configuration Error: {e}")
            except Exception as e:
                st.error(
                    f"❌ API Call Failed: {e}\n\n"
                    "Troubleshooting:\n"
                    "- Verify your API Key is correct\n"
                    "- Check that the Base URL is reachable\n"
                    "- Confirm the model name exists\n"
                    "- Ensure network connectivity is stable"
                )

    # Show cached results when returning to this page
    if "page3_intro" in st.session_state and not generate_btn:
        st.markdown("### Generated Introduction (cached)")
        st.markdown(st.session_state["page3_intro"])




# ==================== Page 4: About ACTA ====================
elif page == "ℹ️ About ACTA":
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    # Load content from the about Word document or Markdown file
    about_content = load_prompt_from_docx(
        PROMPT_FILE_PAGE4, DEFAULT_PROMPT_PAGE4
    )

    # Render as Markdown for proper formatting
    st.markdown(about_content)

    st.markdown("<hr>", unsafe_allow_html=True)

    # Display QR code image if present in project directory
    qr_paths = []
    for fname in ["QRcode.png", "QRcode.jpg"]:
        candidate = os.path.join(APP_DIR, fname)
        if os.path.exists(candidate):
            qr_paths.append(candidate)

    if qr_paths:
        st.markdown(
            '<p style="text-align:center; color:#64748b; '
            'font-size:13px; margin-top:32px;">— — —</p>',
            unsafe_allow_html=True,
        )
        for qr_path in qr_paths:
            st.image(qr_path, width=180)
    else:
        st.markdown(
            '<p style="text-align:center; color:#94a3b8; '
            'font-size:12px; margin-top:24px;">'
            '(Place QRcode.png or QRcode.jpg in the project folder to display here)'
            '</p>',
            unsafe_allow_html=True,
        )



# ==================== Page 2: Framework Architect ====================
elif page == "📐 Framework Architect":
    # Prevent the Streamlit toolbar from overlapping the hero title
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        '<div class="page-hero">'
        "📐 Framework Architect</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-tagline">'
        "Provide your paper details and AI will generate a rigorous, "
        "hierarchically structured academic outline."
        "</div>",
        unsafe_allow_html=True,
    )

    # Input area
    paper_title = st.text_input(
        "Paper Title",
        placeholder="e.g. Multimodal Sentiment Analysis with Deep Learning",
        help="Enter the full title of your paper",
        key="paper_title_input",
    )

    paper_description = st.text_area(
        "Brief Description",
        placeholder=(
            "Provide up to 150 words describing your research scope, "
            "methodology, key contributions, or specific directions "
            "you want the outline to emphasize…"
        ),
        help="Optional context to guide the outline generation (max ~150 words)",
        max_chars=1200,
        height=100,
        key="paper_description_input",
    )

    paper_type = st.selectbox(
        "Paper Type",
        [
            "Modeling",
            "Experimental",
            "Survey",
        ],
        help="Select the research methodology category",
        key="paper_type_select",
    )

    col_left, col_right = st.columns([1, 2])
    with col_left:
        generate_btn = st.button(
            "Generate",
            type="primary",
            use_container_width=True,
        )

    st.markdown("<hr>", unsafe_allow_html=True)

    if generate_btn:
        if not paper_title.strip():
            st.warning("⚠️ Please enter a paper title.")
        else:
            try:
                with st.spinner("ACTA is designing the framework ..."):
                    prompt_template = load_prompt_from_docx(
                        PROMPT_FILE_PAGE2, DEFAULT_PROMPT_PAGE2
                    )
                    desc = paper_description.strip() or (
                        "No additional description provided."
                    )
                    final_prompt = (
                        prompt_template
                        .replace("{title}", paper_title.strip())
                        .replace("{paper_type}", paper_type)
                        .replace("{description}", desc)
                    )

                    raw_result = call_llm(final_prompt)
                    rendered_html = render_framework_markdown(raw_result)
                    st.session_state["page2_framework"] = rendered_html


                    st.markdown("### 📋 Paper Outline")
                    st.markdown(rendered_html, unsafe_allow_html=True)
                    st.success("✅ Framework generated successfully!")

            except ValueError as e:
                st.error(f"❌ Configuration Error: {e}")
            except Exception as e:
                st.error(
                    f"❌ API Call Failed: {e}\n\n"
                    "Troubleshooting:\n"
                    "- Verify your API Key is correct\n"
                    "- Check that the Base URL is reachable\n"
                    "- Confirm the model name exists\n"
                    "- Ensure network connectivity is stable"
                )

    # Show cached results when returning to this page
    if "page2_framework" in st.session_state and not generate_btn:
        st.markdown("### Paper Outline (cached)")
        st.markdown(st.session_state["page2_framework"], unsafe_allow_html=True)
