"""
ACTA — AI-Powered Research Assistant
============================================
A Streamlit web application for scientific question synthesis
and academic paper framework design.

Dependencies: streamlit, openai, python-docx
"""

import warnings
warnings.filterwarnings("ignore", message=".*numexpr.*")
warnings.filterwarnings("ignore", message=".*bottleneck.*")
import streamlit as st
import os
import re
import difflib
from openai import OpenAI
from docx import Document
from openpyxl import load_workbook
from msoffcrypto.format.ooxml import OOXMLFile
import msoffcrypto
import io
import subprocess

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
PROMPT_FILE_PAGE5 = os.path.join(APP_DIR, "prompts_page5.docx")

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
DEFAULT_PROMPT_PAGE5 = (
    "You are an academic publishing advisor. Based on the abstract below, "
    "write a brief analysis (approximately 100 words, in English) covering:\n"
    "1. What broad research field this work belongs to (1-2 disciplines).\n"
    "2. Which research_major_direction might be relevant.\n"
    "3. One practical tip for finding suitable journals in our database, "
    "mentioning specific categories from the available columns: "
    "Category (EN) and research_major_direction (specific research area).\n\n"
    "Keep the tone helpful and concise. Do NOT list specific journal names.\n\n"
    "Abstract: {abstract}"
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


def _git_commit_token_update():
    """Best-effort git add + commit + push of user_keys.xlsx after token update.

    Writes status to git_sync.log in the project directory for debugging.
    """
    import datetime
    log_path = os.path.join(APP_DIR, "git_sync.log")

    def _log(msg):
        try:
            with open(log_path, "a", encoding="utf-8") as lf:
                lf.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
        except Exception:
            pass

    try:
        excel_path = os.path.join(APP_DIR, "user_keys.xlsx")
        repo_dir = APP_DIR

        # Step 1: check git repo
        r = subprocess.run(
            ["git", "rev-parse", "--git-dir"],
            cwd=repo_dir, capture_output=True, text=True, timeout=5
        )
        if r.returncode != 0:
            _log("SKIP: not a git repository")
            return

        # Step 2: git add
        r = subprocess.run(
            ["git", "add", "user_keys.xlsx"],
            cwd=repo_dir, capture_output=True, text=True, timeout=10
        )
        if r.returncode != 0:
            _log(f"FAIL git add: {r.stderr.strip()}")
            return

        # Step 3: check if there are staged changes
        r = subprocess.run(
            ["git", "diff", "--cached", "--quiet", "user_keys.xlsx"],
            cwd=repo_dir, capture_output=True, timeout=10
        )
        if r.returncode == 0:
            _log("SKIP: no changes to commit")
            return

        # Step 4: commit
        msg = f"Update token usage - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        r = subprocess.run(
            ["git", "commit", "-m", msg],
            cwd=repo_dir, capture_output=True, text=True, timeout=10
        )
        if r.returncode != 0:
            _log(f"FAIL git commit: {r.stderr.strip()}")
            return
        _log(f"COMMIT: {msg}")

        # Step 5: push
        r = subprocess.run(
            ["git", "push"],
            cwd=repo_dir, capture_output=True, text=True, timeout=30
        )
        if r.returncode != 0:
            _log(f"FAIL git push (rc={r.returncode}): {r.stderr.strip()[:200]}")
            _log("HINT: Run 'git push' manually or configure credential.helper")
            return
        _log("PUSH: success")

    except subprocess.TimeoutExpired:
        _log("FAIL: git operation timed out")
    except FileNotFoundError:
        _log("FAIL: git command not found (is Git installed?)")
    except Exception as e:
        _log(f"FAIL: unexpected error - {e}")




def load_user_key_credentials(user_key_input):
    """Look up a 10-digit user key in the password-protected Excel.

    Columns: user_key | api_key | base_url | model | token_limit | token_used
    Returns (api_key, base_url, model, token_limit, token_used, row_index)
    or (None,)*7 if not found.
    """
    excel_path = os.path.join(APP_DIR, "user_keys.xlsx")
    if not os.path.exists(excel_path):
        return None, None, None, None, None, None

    user_key_input = str(user_key_input).strip()
    if len(user_key_input) != 10 or not user_key_input.isdigit():
        return None, None, None, None, None, None

    EXCEL_PASSWORD = "980120"

    try:
        decrypted = io.BytesIO()
        with open(excel_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=EXCEL_PASSWORD)
            office_file.decrypt(decrypted)
        decrypted.seek(0)

        wb = load_workbook(decrypted, read_only=True)
        ws = wb.active
        row_idx = 2  # 1-based, skip header
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                row_idx += 1
                continue
            stored_key = str(row[0]).strip()
            if stored_key == user_key_input:
                api_key = str(row[1]).strip() if row[1] else None
                base_url = str(row[2]).strip() if row[2] else None
                model = str(row[3]).strip() if len(row) > 3 and row[3] else None
                token_limit = int(row[4]) if len(row) > 4 and row[4] is not None else 200000
                token_used = int(row[5]) if len(row) > 5 and row[5] is not None else 0
                wb.close()
                return api_key, base_url, model, token_limit, token_used, row_idx
            row_idx += 1
        wb.close()
    except Exception:
        pass
    return None, None, None, None, None, None


def update_token_usage(user_key_input, row_index, new_token_used):
    """Write the updated token_used back to the encrypted Excel for a given key row."""
    excel_path = os.path.join(APP_DIR, "user_keys.xlsx")
    EXCEL_PASSWORD = "980120"

    try:
        # 1. Decrypt
        decrypted = io.BytesIO()
        with open(excel_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=EXCEL_PASSWORD)
            office_file.decrypt(decrypted)
        decrypted.seek(0)

        # 2. Load (not read_only) and update
        wb = load_workbook(decrypted)
        ws = wb.active
        ws.cell(row=row_index, column=6, value=new_token_used)  # column 6 = token_used

        # 3. Save to BytesIO
        saved = io.BytesIO()
        wb.save(saved)
        wb.close()
        saved.seek(0)

        # 4. Re-encrypt and write to disk
        with open(excel_path, "wb") as f:
            encrypted_out = io.BytesIO()
            of_enc = OOXMLFile(saved)
            of_enc.encrypt(EXCEL_PASSWORD, encrypted_out)
            encrypted_out.seek(0)
            f.write(encrypted_out.read())
    except Exception as e:
        # Log write-back failures
        try:
            log_path = os.path.join(APP_DIR, "token_sync.log")
            with open(log_path, "a", encoding="utf-8") as lf:
                import datetime
                lf.write(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] FAIL update_token_usage: {e}\n")
        except Exception:
            pass


def call_llm(prompt, system_prompt=SYSTEM_PROMPT):
    """Invoke the LLM API with quota check and token tracking.

    Reads API credentials from st.session_state. Supports User Key override
    with token quota enforcement. Raises ValueError when configuration is
    missing or quota is exceeded.
    """
    # API Key takes priority; User Key is fallback
    api_key = st.session_state.get("api_key", "").strip()
    base_url = st.session_state.get("base_url", "").strip()
    model = st.session_state.get("model", "gpt-3.5-turbo").strip()
    token_limit = None
    token_used = None
    row_index = None
    using_user_key = False

    # If API Key is not set, try User Key
    if not api_key:
        user_key = st.session_state.get("user_key", "").strip()
        if user_key:
            uk_api_key, uk_base_url, uk_model, token_limit, token_used, row_index = (
                load_user_key_credentials(user_key)
            )
            if uk_api_key:
                api_key = uk_api_key
                base_url = uk_base_url or "https://api.siliconflow.cn/v1"
                model = uk_model or "deepseek-ai/DeepSeek-V4-Pro"
                using_user_key = True

                # Quota check
                if token_limit is not None and token_used is not None:
                    if token_used >= token_limit:
                        raise ValueError(
                            f"Token quota exhausted ({token_used:,}/{token_limit:,} tokens used). "
                            "Please contact the administrator to increase your quota."
                        )
            else:
                raise ValueError(
                    "Invalid User Key. Please check your 10-digit key or contact the administrator."
                )

    if not api_key:
        raise ValueError("Please configure your API Key or User Key in the sidebar.")
    if not base_url:
        raise ValueError("Please configure the API Base URL in the sidebar.")
    if not model:
        raise ValueError("Please configure the Model Name in the sidebar.")

    # Create HTTP client with timeout
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
        timeout=httpx.Timeout(120.0, connect=60.0),
    )

    # Update token usage for user-key users
    if using_user_key and row_index is not None and token_used is not None:
        try:
            usage = response.usage
            if usage and usage.total_tokens:
                new_total = token_used + usage.total_tokens
                update_token_usage(user_key, row_index, new_total)
                # Update sidebar display
                st.session_state["user_token_used"] = new_total
                # Git auto-commit (best-effort, non-blocking)
                _git_commit_token_update()
        except Exception:
            pass

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


# ==================== Journal Recommender Helpers ====================

JOURNAL_HEADERS = [
    "journal_id", "journal_name", "impact_factor", "citescore",
    "xr_zone", "top", "category_en", "category",
    "research_major_direction", "research_minor_direction",
    "is_oa", "language", "publisher", "self_citation_rate",
    "five_year_if", "h_index", "journal_intro", "website",
    "research_direction", "country", "period", "pub_year",
    "gold_oa_ratio", "research_article_ratio", "review_speed",
    "acceptance_rate", "annual_articles", "citescore_detail",
    "p_issn", "e_issn"
]

JOURNAL_DISPLAY_LABELS = {
    "journal_id": "Journal ID",
    "journal_name": "Journal Name",
    "impact_factor": "Impact Factor",
    "citescore": "CiteScore",
    "xr_zone": "New Sharp Zone",
    "top": "TOP",
    "category_en": "Category (EN)",
    "category": "Category",
    "research_major_direction": "Major Research Direction",
    "research_minor_direction": "Minor Research Direction",
    "is_oa": "Open Access",
    "language": "Language",
    "publisher": "Publisher",
    "self_citation_rate": "Self-citation Rate",
    "five_year_if": "5-Year IF",
    "h_index": "H-index",
    "journal_intro": "Introduction",
    "website": "Website",
    "research_direction": "Research Direction",
    "country": "Country",
    "period": "Period",
    "pub_year": "Publication Year",
    "gold_oa_ratio": "Gold OA Ratio",
    "research_article_ratio": "Research Article Ratio",
    "review_speed": "Review Speed",
    "acceptance_rate": "Acceptance Rate",
    "annual_articles": "Annual Articles",
    "citescore_detail": "CiteScore Detail",
    "p_issn": "P-ISSN",
    "e_issn": "E-ISSN",
}

FILTER_COLUMNS = {"category", "research_major_direction", "impact_factor", "citescore", "xr_zone", "top"}
HEADER_COLUMNS = {"journal_name", "impact_factor", "xr_zone"}
DISPLAY_COLUMNS = [h for h in JOURNAL_HEADERS if h not in HEADER_COLUMNS]


def load_journal_data():
    jdata_path = os.path.join(APP_DIR, "Jdata_with_cn_v1.xlsx")
    wb = load_workbook(jdata_path, read_only=True)
    ws = wb.active
    journals = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[8] is not None and str(row[8]).strip() != "":
            d = {h: (row[i] if i < len(row) else None) for i, h in enumerate(JOURNAL_HEADERS)}
            journals.append(d)
    wb.close()
    return journals


def get_unique_values(journals, col):
    """Get sorted unique non-null values for a column."""
    vals = set()
    for j in journals:
        v = j.get(col)
        if v is not None and str(v).strip() != "" and str(v).strip().lower() != "none":
            vals.add(str(v).strip())
    try:
        return sorted(vals, key=lambda x: (
            float(x) if x.replace(".", "").replace("-", "").lstrip("-").isdigit() else float("inf"), x
        ))
    except Exception:
        return sorted(vals)


def filter_journals_by_selection(journals, selections, skip_key=None):
    """Filter journals by selected values, optionally skipping one filter key."""
    filtered = []
    for journal in journals:
        keep = True
        for key, selected_values in selections.items():
            if key == skip_key or not selected_values:
                continue
            journal_value = str(journal.get(key) or "").strip()
            if journal_value not in selected_values:
                keep = False
                break
        if keep:
            filtered.append(journal)
    return filtered


def normalize_filter_values(raw_value):
    """Normalize a filter value into a clean list of strings."""
    if isinstance(raw_value, str):
        raw_value = [raw_value] if raw_value else []
    elif raw_value is None:
        raw_value = []
    return [str(v).strip() for v in raw_value if str(v).strip()]


def build_multiselect_options(options, selected_values):
    """Keep selected values visible while showing current available options."""
    merged = list(selected_values)
    for option in options:
        if option not in merged:
            merged.append(option)
    return merged


def find_journal_matches(journals, query, limit=5):
    """Return up to `limit` closest journal matches for a free-text query."""
    query_norm = str(query or "").strip().lower()
    if not query_norm:
        return []

    exact_matches = []
    ranked_matches = []
    for journal in journals:
        name = str(journal.get("journal_name") or "").strip()
        if not name:
            continue
        name_norm = name.lower()
        if name_norm == query_norm:
            exact_matches.append(journal)
            continue
        score = difflib.SequenceMatcher(None, query_norm, name_norm).ratio()
        if query_norm in name_norm:
            score += 0.35
        if name_norm.startswith(query_norm):
            score += 0.25
        if score > 0.2:
            ranked_matches.append((score, name, journal))

    if exact_matches:
        return exact_matches[:1]

    ranked_matches.sort(key=lambda item: (-item[0], item[1].lower()))
    return [journal for _, _, journal in ranked_matches[:limit]]


def render_journal_entries(journals):
    """Render journal rows as expandable entries."""
    for j in journals:
        name = str(j.get("journal_name") or "Unknown")
        impact = j.get("impact_factor")
        if_str = ("%.2f" % impact) if impact is not None else "N/A"
        xr = j.get("xr_zone")
        xr_str = (" | 新锐分区 " + str(xr)) if xr is not None else ""
        label = name + " (IF: " + if_str + ")" + xr_str
        with st.expander(label):
            for col_key in DISPLAY_COLUMNS:
                val = j.get(col_key)
                if val is not None and str(val).strip() != "" and str(val).strip().lower() != "none":
                    lbl = JOURNAL_DISPLAY_LABELS.get(col_key, col_key)
                    if col_key == "website" and val:
                        st.markdown("**" + lbl + ":** [" + str(val) + "](" + str(val) + ")")
                    else:
                        st.markdown("**" + lbl + ":** " + str(val))


def prune_journal_filter_chain(journals, filter_specs, selections):
    """Prune staged filter selections and compute downstream options."""
    pruned = {}
    options_map = {}
    current_pool = journals
    for data_key, _, _ in filter_specs:
        options = get_unique_values(current_pool, data_key)
        valid_selected = [value for value in selections.get(data_key, []) if value in options]
        pruned[data_key] = valid_selected
        options_map[data_key] = build_multiselect_options(options, valid_selected)
        if valid_selected:
            current_pool = filter_journals_by_selection(current_pool, {data_key: valid_selected})
    return pruned, options_map


def apply_journal_filter_stage(stage_key):
    """Commit one filter stage and prune only downstream stages."""
    filter_specs = st.session_state.get("page5_filter_specs", [])
    journals_all = st.session_state.get("cached_journals", [])
    if not filter_specs or not journals_all:
        return

    current_filters = st.session_state.get("page5_filters", {})
    selections = {
        data_key: normalize_filter_values(current_filters.get(data_key, []))
        for data_key, _, _ in filter_specs
    }
    widget_key = "page5_" + stage_key + "_widget"
    selections[stage_key] = normalize_filter_values(st.session_state.get(widget_key, []))

    pruned, _ = prune_journal_filter_chain(journals_all, filter_specs, selections)
    st.session_state["page5_filters"] = pruned
    for data_key, _, _ in filter_specs:
        st.session_state["page5_" + data_key + "_widget"] = list(pruned[data_key])
    st.session_state["page5_page"] = 1

def render_journal_results():
    "Render journal list with linked multi-select filters and expandable rows."
    journals_all = st.session_state.get("cached_journals", [])
    if not journals_all:
        st.info("Loading journal database ...")
        with st.spinner("Loading journals ..."):
            st.session_state["cached_journals"] = load_journal_data()
        st.rerun()
    if not journals_all:
        st.info("Journal database not loaded. Please reload the page.")
        return

    filter_specs = [
        ("category", "Category", "page5_cat"),
        ("research_major_direction", "Research Major Direction", "page5_major"),
        ("xr_zone", "新锐分区", "page5_xr"),
        ("top", "TOP", "page5_top"),
    ]
    st.session_state["page5_filter_specs"] = filter_specs
    if "page5_filters" not in st.session_state:
        st.session_state["page5_filters"] = {data_key: [] for data_key, _, _ in filter_specs}

    stored_filters = st.session_state["page5_filters"]
    selections = {data_key: normalize_filter_values(stored_filters.get(data_key, [])) for data_key, _, _ in filter_specs}

    title_col, mode_col = st.columns([2.3, 1.7])
    with title_col:
        st.markdown("### Filter Journals")
    with mode_col:
        search_mode = st.radio(
            "Search Mode",
            ["Category Search", "Journal Search"],
            key="page5_search_mode",
            horizontal=True,
            label_visibility="collapsed",
        )

    if search_mode == "Journal Search":
        journal_query = st.text_input(
            "Journal Name",
            key="page5_journal_query",
            placeholder="Type a journal name",
        )
        if not str(journal_query or "").strip():
            st.info("Enter a journal name to search for the closest match.")
            return

        journal_matches = find_journal_matches(journals_all, journal_query, limit=5)
        if not journal_matches:
            st.info("No journal match found.")
            return

        st.markdown("**Showing " + str(len(journal_matches)) + " journal match(es)**")
        render_journal_entries(journal_matches)
        return

    applied_selections, filter_options = prune_journal_filter_chain(journals_all, filter_specs, selections)
    if applied_selections != selections:
        st.session_state["page5_filters"] = applied_selections
        selections = applied_selections
    else:
        selections = applied_selections

    for data_key, _, _ in filter_specs:
        widget_key = "page5_" + data_key + "_widget"
        if widget_key not in st.session_state:
            st.session_state[widget_key] = list(selections[data_key])
        elif normalize_filter_values(st.session_state.get(widget_key, [])) != selections[data_key]:
            st.session_state[widget_key] = list(selections[data_key])

    c1, c2 = st.columns(2)
    with c1:
        sel_cat = st.multiselect(
            "Category",
            filter_options["category"],
            key="page5_category_widget",
            on_change=apply_journal_filter_stage,
            args=("category",),
        )
    with c2:
        sel_major = st.multiselect(
            "Research Major Direction",
            filter_options["research_major_direction"],
            key="page5_research_major_direction_widget",
            on_change=apply_journal_filter_stage,
            args=("research_major_direction",),
        )

    c3, c4, c5 = st.columns(3)
    with c3:
        sel_xr = st.multiselect(
            "新锐分区",
            filter_options["xr_zone"],
            key="page5_xr_zone_widget",
            on_change=apply_journal_filter_stage,
            args=("xr_zone",),
        )
    with c4:
        sel_top = st.multiselect(
            "TOP",
            filter_options["top"],
            key="page5_top_widget",
            on_change=apply_journal_filter_stage,
            args=("top",),
        )
    with c5:
        sort_options = ["IF: High to Low", "IF: Low to High", "CiteScore: High to Low", "CiteScore: Low to High"]
        sel_sort = st.selectbox("Sort By", sort_options, key="page5_sort")

    sel_cat = selections["category"]
    sel_major = selections["research_major_direction"]
    sel_xr = selections["xr_zone"]
    sel_top = selections["top"]

    current_signature = (
        tuple(sorted(sel_cat)),
        tuple(sorted(sel_major)),
        tuple(sorted(sel_xr)),
        tuple(sorted(sel_top)),
    )
    if st.session_state.get("page5_filter_signature") != current_signature:
        st.session_state["page5_filter_signature"] = current_signature
        st.session_state["page5_page"] = 1

    filtered = filter_journals_by_selection(
        journals_all,
        {
            "category": sel_cat,
            "research_major_direction": sel_major,
            "xr_zone": sel_xr,
            "top": sel_top,
        },
    )

    if "CiteScore" in sel_sort:
        if "High to Low" in sel_sort:
            filtered.sort(key=lambda x: -(x.get("citescore") or 0))
        else:
            filtered.sort(key=lambda x: (x.get("citescore") or 999999))
    else:
        if "High to Low" in sel_sort:
            filtered.sort(key=lambda x: -(x.get("impact_factor") or 0))
        else:
            filtered.sort(key=lambda x: (x.get("impact_factor") or 999999))

    PAGE_SIZE = 20
    total_pages = max(1, (len(filtered) + PAGE_SIZE - 1) // PAGE_SIZE)
    if "page5_page" not in st.session_state:
        st.session_state["page5_page"] = 1
    pg = st.session_state["page5_page"]
    if pg > total_pages:
        pg = total_pages
        st.session_state["page5_page"] = pg
    start_idx = (pg - 1) * PAGE_SIZE
    end_idx = min(start_idx + PAGE_SIZE, len(filtered))

    st.markdown("**Showing " + str(len(filtered)) + " journal(s) (Page " + str(pg) + " of " + str(total_pages) + ")**")

    if not filtered:
        st.info("No journals match the current filters. Try adjusting your selection.")
        return

    render_journal_entries(filtered[start_idx:end_idx])

    if total_pages > 1:
        pc1, pc2, pc3, pc4, pc5 = st.columns([1, 2, 1, 2, 1])
        with pc2:
            if st.button("< Previous", disabled=(pg <= 1), key="pg5_prev", use_container_width=True):
                st.session_state["page5_page"] = max(1, pg - 1)
                st.rerun()
        with pc4:
            if st.button("Next >", disabled=(pg >= total_pages), key="pg5_next", use_container_width=True):
                st.session_state["page5_page"] = min(total_pages, pg + 1)
                st.rerun()

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
            "📋 Journal Recommender",
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
        user_key = st.text_input(
            "User Key",
            type="password",
            value=st.session_state.get("user_key", ""),
            placeholder="10-digit key (if provided)",
            help="Enter a 10-digit user key to use the internal API. Leave blank to use your own API Key below.",
        )
        st.session_state["user_key"] = user_key

        # Validate user key and show token status
        if user_key and len(str(user_key).strip()) == 10 and str(user_key).strip().isdigit():
            uk_api, uk_url, uk_model, tok_limit, tok_used, _ = load_user_key_credentials(user_key)
            if uk_api and tok_limit is not None:
                st.session_state["user_key_valid"] = True
                st.session_state["user_token_limit"] = tok_limit
                st.session_state["user_token_used"] = tok_used
                remaining = max(0, tok_limit - tok_used)
                pct = (tok_used / tok_limit * 100) if tok_limit > 0 else 0
                if pct >= 90:
                    color = "#ef4444"
                elif pct >= 50:
                    color = "#f59e0b"
                else:
                    color = "#22c55e"
                st.markdown(
                    f'<p style="font-size:11px; color:{color}; margin:-8px 0 8px 0;">'
                    f"Remaining: {remaining:,} / {tok_limit:,} tokens ({100-pct:.0f}%)"
                    f"</p>",
                    unsafe_allow_html=True,
                )
            else:
                st.session_state["user_key_valid"] = False
                st.markdown(
                    '<p style="font-size:11px; color:#ef4444; margin:-8px 0 8px 0;">'
                    "Invalid User Key</p>",
                    unsafe_allow_html=True,
                )
        elif user_key:
            st.session_state["user_key_valid"] = False
            st.markdown(
                '<p style="font-size:11px; color:#ef4444; margin:-8px 0 8px 0;">'
                "Key must be 10 digits</p>",
                unsafe_allow_html=True,
            )


        # Check git sync status
        git_log_path = os.path.join(APP_DIR, "git_sync.log")
        git_warn = False
        if os.path.exists(git_log_path):
            try:
                with open(git_log_path, "r", encoding="utf-8") as lf:
                    lines = lf.readlines()
                    if lines:
                        last = lines[-1]
                        if "FAIL" in last:
                            git_warn = True
                            st.warning(
                                "GitHub sync may not be working. "
                                "Check git_sync.log in the project folder."
                            )
            except Exception:
                pass

        st.markdown(
            '<p style="font-size:11px; color:#94a3b8; margin:-10px 0 10px 0;">'
            "— OR —"
            "</p>",
            unsafe_allow_html=True,
        )

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


# ==================== Page 5: Journal Recommender ====================
if page == "\U0001f4cb Journal Recommender":
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        '<div class="page-hero">\U0001f4cb Journal Recommender</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="page-tagline">Browse and filter journals by discipline, impact factor, and zone.</div>',
        unsafe_allow_html=True,
    )

    abstract = st.text_area(
        "Paper Abstract / Topic Description",
        placeholder="Paste your paper abstract or describe your research topic. ACTA will analyze it and suggest relevant fields for journal selection.",
        height=150,
        key="journal_abstract_input",
    )

    col_left, col_right = st.columns([1, 2])
    with col_left:
        analyze_btn = st.button("Analyze", type="primary", use_container_width=True, key="journal_analyze_btn")

    if analyze_btn:
        if not abstract.strip():
            st.warning("Please enter a paper abstract.")
        else:
            try:
                with st.spinner("ACTA is analyzing your research ..."):
                    prompt_template = load_prompt_from_docx(PROMPT_FILE_PAGE5, DEFAULT_PROMPT_PAGE5)
                    final_prompt = prompt_template.replace("{abstract}", abstract.strip())
                    advice = call_llm(final_prompt)
                    st.session_state["page5_advice"] = advice
                    st.session_state["page5_abstract"] = abstract.strip()
            except ValueError as e:
                st.error("Configuration Error: " + str(e))
            except Exception as e:
                st.error("Error: " + str(e))

    if "page5_advice" in st.session_state:
        st.markdown("---")
        st.markdown("### Field Analysis and Advice")
        st.info(st.session_state["page5_advice"])

    st.markdown("---")

    render_journal_results()

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
        help="Provide a comprehensive description of your research topic (max ~300 words)",
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
