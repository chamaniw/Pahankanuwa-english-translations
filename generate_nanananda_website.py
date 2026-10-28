import os
import shutil
import re
from pathlib import Path
import pandas as pd
import unicodedata
import json

# Optional import for reading DOCX files
try:
    import docx
except ImportError:
    docx = None

# ==========================================
# PATH & DIRECTORY CONFIGURATION
# ==========================================
BASE_DIR = Path(r"C:\Pahankanuwa")
DOCS_DIR = BASE_DIR / "docs"
DOCS_DOCX_DIR = DOCS_DIR / "docx"
DOCS_PDF_DIR = DOCS_DIR / "pdf"
DOCS_EPUB_DIR = DOCS_DIR / "epub"
DOCS_AUDIO_DIR = DOCS_DIR / "audio"

# Ensure all destination output directories exist
DOCS_DIR.mkdir(parents=True, exist_ok=True)
DOCS_DOCX_DIR.mkdir(parents=True, exist_ok=True)
DOCS_PDF_DIR.mkdir(parents=True, exist_ok=True)
DOCS_EPUB_DIR.mkdir(parents=True, exist_ok=True)
DOCS_AUDIO_DIR.mkdir(parents=True, exist_ok=True)

# Source asset folders
SRC_DOCX_DIR = BASE_DIR / "docx"
SRC_PDF_DIR = BASE_DIR / "pdf"
SRC_EPUB_DIR = BASE_DIR / "epub"
SRC_AUDIO_DIR = BASE_DIR / "audio"
SRC_TXT_DIR = BASE_DIR / "txt"  # Fallback TXT directory if present

# Complete collection file names
DOCX_COMPLETE = "Pahankanuwa_Sermons_Complete.docx"
PDF_COMPLETE = "Pahankanuwa_Sermons_Complete.pdf"
EPUB_COMPLETE = "Pahankanuwa_Sermons_Complete.epub"
ZIP_FILE = "Pahankanuwa_English_Translations.zip"


# ==========================================
# DOCX / TXT CONTENT EXTRACTOR
# ==========================================
def read_sermon_content(s_num):
    """
    Reads the content of a sermon from DOCX or TXT files and returns clean HTML paragraphs.
    """
    docx_filename = find_sermon_file(s_num, "docx", SRC_DOCX_DIR)
    if docx_filename and (SRC_DOCX_DIR / docx_filename).exists():
        docx_path = SRC_DOCX_DIR / docx_filename
        if docx is not None:
            try:
                doc = docx.Document(docx_path)
                html_paragraphs = []
                for p in doc.paragraphs:
                    text = p.text.strip()
                    if text:
                        formatted_text = text.replace("\n", "<br>")
                        html_paragraphs.append(f"<p>{formatted_text}</p>")
                if html_paragraphs:
                    return "\n".join(html_paragraphs)
            except Exception as e:
                print(f"Error reading DOCX file {docx_path}: {e}")

    txt_filename = find_sermon_file(s_num, "txt", SRC_TXT_DIR)
    if txt_filename and (SRC_TXT_DIR / txt_filename).exists():
        try:
            with open(SRC_TXT_DIR / txt_filename, "r", encoding="utf-8") as f:
                content = f.read().strip()
                paras = [f"<p>{p.strip().replace('\n', '<br>')}</p>" for p in content.split("\n\n") if p.strip()]
                return "\n".join(paras)
        except Exception as e:
            print(f"Error reading TXT file: {e}")

    return f"<p><em>Text content for Sermon {s_num} is currently being prepared. You can download the source document using the links below.</em></p>"


# ==========================================
# HELPER FUNCTIONS
# ==========================================
def parse_and_link_sermons(val):
    if pd.isna(val) or not str(val).strip() or str(val).lower() == "nan":
        return ""
    val_str = str(val).strip()
    if val_str.isdigit() and len(val_str) <= 3:
        s_num = val_str.zfill(3)
        doc_filename = f"sermon_{s_num}_ENGLISH.docx"
        return f'<a class="doc-btn" href="docx/{doc_filename}" target="_blank">📄 Sermon {s_num}</a>'

    items = re.split(r'[,;]', val_str)
    links = []
    for item in items:
        item_str = item.strip()
        match = re.search(r'(\d{1,3})', item_str)
        if match:
            s_num = match.group(1).zfill(3)
            doc_filename = f"sermon_{s_num}_ENGLISH.docx"
            links.append(f'<a class="doc-btn" href="docx/{doc_filename}" target="_blank">📄 Sermon {s_num}</a>')
        else:
            if item_str:
                links.append(item_str)
    return " ".join(links) if links else val_str


def find_sermon_file(s_num, format_ext, search_dir):
    if not search_dir.exists():
        return None
    pattern = re.compile(rf".*{s_num}.*\.{format_ext}$", re.IGNORECASE)
    for file in search_dir.iterdir():
        if file.is_file() and pattern.match(file.name):
            return file.name
    return None

def normalize_text(text):
    if not text:
        return ""

    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(
        ch for ch in text
        if unicodedata.category(ch) != "Mn"
    )
    return text.lower()


# ==========================================
# DATA LOADING & PROCESSOR
# ==========================================
def load_and_process_excel_data():
    scratch_file = BASE_DIR / "pahankanuwa_e_from_scratch_6.xlsx"
    idx_file = BASE_DIR / "Nanananda_Dhamma_Index_V3.xlsx"
    kb_file = BASE_DIR / "Nanananda_Knowledge_Base_V1.xlsx"

    datasets = {
        "sermons": pd.DataFrame(),
        "citations": pd.DataFrame(),
        "statistics": pd.DataFrame()
    }

    if scratch_file.exists():
        df_scratch = pd.read_excel(scratch_file)
        sermons_list = []
        for _, row in df_scratch.iterrows():
            fn = str(row['FileName']).strip() if 'FileName' in row else ""
            if not fn or fn.lower() == 'nan':
                continue
            
            num_match = re.search(r'(\d{1,3})', fn)
            s_num = num_match.group(1).zfill(3) if num_match else "000"
            clean_title = f"Pahankanuwa Sermon {s_num}"

            docx_file = find_sermon_file(s_num, "docx", SRC_DOCX_DIR) or f"sermon_{s_num}_ENGLISH.docx"
            pdf_file = find_sermon_file(s_num, "pdf", SRC_PDF_DIR) or f"sermon_{s_num}.pdf"
            epub_file = find_sermon_file(s_num, "epub", SRC_EPUB_DIR) or f"sermon_{s_num}.epub"

            para_7 = str(row['Para_7']).strip() if 'Para_7' in row and pd.notna(row['Para_7']) else ""
            excerpt = para_7 if para_7 and para_7.lower() != "nan" else "No excerpt available"

            para_9 = str(row['Para_9']).strip() if 'Para_9' in row and pd.notna(row['Para_9']) else ""
            english_translation = para_9 if para_9 and para_9.lower() != "nan" else "No English translation available"

            para_11 = str(row['Para_11']).strip() if 'Para_11' in row and pd.notna(row['Para_11']) else ""
            likely_source = para_11 if para_11 and para_11.lower() != "nan" else "Unspecified Canonical Source"

            doc_links = (
                f'<a class="doc-btn" href="docx/{docx_file}" target="_blank">📄 DOCX</a> '
                f'<a class="doc-btn" href="pdf/{pdf_file}" target="_blank">📕 PDF</a> '
                f'<a class="doc-btn" href="epub/{epub_file}" target="_blank">📘 EPUB</a>'
            )

            sermon_link_html = f'<a href="sermon_{s_num}.html" class="sermon-title-link"><strong>{clean_title}</strong></a>'

            sermons_list.append({
                "Sermon": sermon_link_html,
                "Opening Passage Excerpt": excerpt.replace('\n', '<br>'),
                "English Translation": english_translation.replace('\n', '<br>'),
                "Likely Source": likely_source.replace('\n', '<br>'),
                "Document Link": doc_links
            })
        datasets["sermons"] = pd.DataFrame(sermons_list)

    if idx_file.exists():
        xls_idx = pd.ExcelFile(idx_file)
        if 'Pali_Citations' in xls_idx.sheet_names:
            df_cit = pd.read_excel(xls_idx, sheet_name='Pali_Citations')
            df_cit['Sermon Link'] = df_cit['Sermon'].apply(parse_and_link_sermons)
            datasets["citations"] = df_cit[['Citation', 'Sermon Link', 'Type', 'Likely Source']].drop_duplicates()

    if kb_file.exists():
        xls_kb = pd.ExcelFile(kb_file)
        if 'Statistics' in xls_kb.sheet_names:
            datasets["statistics"] = pd.read_excel(xls_kb, sheet_name='Statistics')

    metrics = {
        "sermons_count": len(datasets["sermons"]) if not datasets["sermons"].empty else 202,
        "citations_count": len(datasets["citations"]) if not datasets["citations"].empty else 2815,
    }

    return metrics, datasets


# ==========================================
# GLOBAL CSS STYLING
# ==========================================
SITE_CSS = """
:root {
    --header-bg: #800000;
    --header-text: #ffffff;
    --header-subtitle: #ffd1d1;
    --primary-color: #5c2c16;
    --accent-color: #c88a2a;
    --bg-color: #fdfbf7;
    --card-bg: #ffffff;
    --text-color: #2c2c2c;
    --text-muted: #666666;
    --border-color: #e2d7c7;
    --table-header-bg: #800000;
}

* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    background-color: var(--bg-color);
    color: var(--text-color);
    line-height: 1.6;
}

.site-header {
    background-color: var(--header-bg);
    color: var(--header-text);
    padding: 1.2rem 2rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
.header-container {
    max-width: 1200px;
    margin: 0 auto;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 1rem;
}
.brand-section .logo-title {
    font-family: 'Georgia', serif;
    font-size: 1.8rem;
    font-weight: bold;
    color: #ffffff;
}
.brand-section .subtitle {
    font-size: 0.95rem;
    color: var(--header-subtitle);
    font-style: italic;
}

.top-nav-wrapper {
    display: flex;
    align-items: center;
    gap: 1rem;
    flex-wrap: wrap;
}

.top-nav {
    display: flex;
    gap: 0.5rem;
    flex-wrap: wrap;
    align-items: center;
}
.top-nav a {
    color: #ffffff;
    text-decoration: none;
    padding: 0.5rem 0.9rem;
    border-radius: 4px;
    font-weight: 500;
    font-size: 0.95rem;
    transition: background-color 0.2s, color 0.2s;
}
.top-nav a:hover, .top-nav a.active {
    background-color: rgba(255, 255, 255, 0.25);
    color: #ffffff;
}

.header-search-form {
    display: flex;
    align-items: center;
}
.header-search-input {
    padding: 0.45rem 0.8rem;
    font-size: 0.88rem;
    border: 1px solid rgba(255, 255, 255, 0.4);
    border-radius: 4px;
    background-color: rgba(255, 255, 255, 0.15);
    color: #ffffff;
    outline: none;
    transition: background-color 0.2s, border-color 0.2s;
}
.header-search-input::placeholder {
    color: #ffd1d1;
}
.header-search-input:focus {
    background-color: #ffffff;
    color: #2c2c2c;
    border-color: #ffffff;
}

.search-container {
    max-width: 1200px;
    margin: 1.5rem auto 0 auto;
    padding: 0 1.5rem;
}
.search-input {
    width: 100%;
    padding: 0.85rem 1.2rem;
    font-size: 1rem;
    border: 1px solid var(--border-color);
    border-radius: 6px;
    background-color: #ffffff;
    box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    outline: none;
}
.search-input:focus {
    border-color: var(--header-bg);
    box-shadow: 0 0 0 3px rgba(128, 0, 0, 0.1);
}

.homepage-content {
    max-width: 1200px;
    margin: 1rem auto 3rem auto;
    padding: 0 0.5rem;
}

.hero {
    background: var(--card-bg);
    padding: 2rem;
    border-radius: 8px;
    border: 1px solid var(--border-color);
    margin-bottom: 2rem;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
}
.hero h1 { font-family: 'Georgia', serif; font-size: 1.8rem; color: #222; margin-bottom: 0.3rem; }
.hero h2 { font-family: 'Georgia', serif; font-size: 1.2rem; color: var(--primary-color); margin-bottom: 1.2rem; font-weight: normal; }
.hero p { margin-bottom: 1rem; color: #333; font-size: 0.98rem; text-align: left; }
.pali-quote { font-style: italic; color: var(--header-bg); text-align: center; margin: 1.5rem 0; font-weight: 500; }
.hero p.disclaimer { font-size: 0.88rem; color: var(--text-muted); font-style: italic; border-top: 1px solid var(--border-color); padding-top: 1rem; margin-top: 1.5rem; }

.download-banner {
    display: flex;
    justify-content: center;
    gap: 0.75rem;
    flex-wrap: wrap;
    margin: 1.5rem 0;
}

.btn {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.6rem 1.2rem;
    background-color: var(--header-bg);
    color: white !important;
    text-decoration: none;
    border-radius: 6px;
    font-weight: 600;
    font-size: 0.9rem;
    transition: background-color 0.2s;
}
.btn:hover { background-color: #5c0000; }
.btn-secondary {
    background-color: transparent;
    color: var(--header-bg) !important;
    border: 1px solid var(--header-bg);
}
.btn-secondary:hover { background-color: #f4eee6; }

.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 1.5rem;
    margin-bottom: 2.5rem;
}
.stat-card {
    background: var(--card-bg);
    border: 1px solid var(--border-color);
    padding: 2rem 1rem;
    text-align: center;
    border-radius: 8px;
    box-shadow: 0 2px 5px rgba(0,0,0,0.02);
    text-decoration: none;
    color: inherit;
    transition: transform 0.15s, border-color 0.15s;
}
.stat-card:hover {
    transform: translateY(-2px);
    border-color: var(--accent-color);
}
.stat-number {
    display: block;
    font-family: 'Georgia', serif;
    font-size: 2.5rem;
    font-weight: bold;
    color: var(--primary-color);
    line-height: 1.1;
}
.stat-label {
    font-size: 0.85rem;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-top: 0.4rem;
}

.download-list {
    display: flex;
    flex-direction: column;
    gap: 1rem;
    margin-top: 1.5rem;
}
.download-item {
    background: var(--card-bg);
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 1.25rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 1rem;
}
.download-info h4 { margin: 0 0 0.25rem 0; color: var(--header-bg); font-family: 'Georgia', serif; }
.download-info p { margin: 0; color: var(--text-muted); font-size: 0.9rem; }
.download-links { display: flex; gap: 0.5rem; flex-wrap: wrap; }

.table-container {
    background: #ffffff;
    border: 1px solid var(--border-color);
    border-radius: 8px;
    overflow-x: auto;
    margin-top: 1.5rem;
    box-shadow: 0 2px 6px rgba(0,0,0,0.02);
}
table.data-table {
    width: 100%;
    border-collapse: collapse;
    text-align: left;
    font-size: 0.92rem;
}
table.data-table th {
    background-color: var(--table-header-bg);
    color: #ffffff;
    padding: 0.8rem 1rem;
    font-weight: 600;
}
table.data-table td {
    padding: 0.8rem 1rem;
    border-bottom: 1px solid var(--border-color);
    vertical-align: top;
}
table.data-table tr:nth-child(even) { background-color: #fcfaf7; }
table.data-table tr:hover { background-color: #f4eee6; }

a.doc-btn {
    display: inline-block;
    background-color: #f4eee6;
    color: var(--header-bg);
    border: 1px solid var(--header-bg);
    padding: 0.35rem 0.75rem;
    border-radius: 4px;
    text-decoration: none;
    font-weight: 600;
    font-size: 0.85rem;
    margin: 2px 2px;
    white-space: nowrap;
    transition: all 0.15s ease-in-out;
}
a.doc-btn:hover {
    background-color: var(--header-bg);
    color: #ffffff;
}

a.sermon-title-link {
    color: var(--header-bg);
    text-decoration: underline;
    text-underline-offset: 3px;
    font-weight: bold;
    transition: color 0.15s ease-in-out;
}
a.sermon-title-link:hover {
    color: var(--accent-color);
}

.sermon-reader {
    background: #ffffff;
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 1rem;
    margin-top: 1.5rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    font-size: 1.05rem;
    line-height: 1.8;
}
@media (min-width: 768px) {
    .sermon-reader {
        padding: 2rem 3rem;
    }
}
.sermon-reader p {
    margin-bottom: 1.25rem;
    text-align: left;
}
.sermon-nav-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 1.5rem;
    padding-bottom: 1rem;
    border-bottom: 1px solid var(--border-color);
}

footer {
    text-align: center;
    padding: 2rem;
    font-size: 0.85rem;
    color: var(--text-muted);
    border-top: 1px solid var(--border-color);
    margin-top: 3rem;
}
"""

SEARCH_JS = """
<script>
function filterData() {
    var input = document.getElementById('searchInput');
    var filter = input.value.toLowerCase();
    
    var rows = document.querySelectorAll('.data-table tbody tr, .download-item');
    rows.forEach(function(row) {
        var text = row.innerText.toLowerCase();
        row.style.display = text.includes(filter) ? '' : 'none';
    });
}
</script>
"""

FOOTER_HTML = f"""
  <footer>
    <p style="margin-bottom: 1rem;">
      <a class="btn" href="{ZIP_FILE}" download>
        📦 Download All Sermons (.ZIP Archive)
      </a>
    </p>
    <p>Ven. Katukurunde Nanananda Thero Dhamma Knowledge Base &copy; 2026.</p>
  </footer>
"""


def build_header_html(active_tab="home"):
    tabs = [
        ("home", "Home", "index.html"),
        ("sermons", "Sermons", "sermons.html"),
        ("downloads", "Downloads", "downloads.html"),
        ("audio", "Audio", "audio.html"),
        ("search", "Search", "search.html"),
        ("statistics", "Statistics", "statistics.html"),
    ]
    
    nav_links = [
        f'<a href="{link}"{" class=\"active\"" if active_tab == tab_id else ""}>{label}</a>'
        for tab_id, label, link in tabs
    ]

    return f"""
  <header class="site-header">
    <div class="header-container">
      <div class="brand-section">
        <div class="logo-title">Ven. Katukurunde Nanananda Thero</div>
        <div class="subtitle">English Translations of Pahankanuwa Sermons</div>
      </div>
      <div class="top-nav-wrapper">
        <nav class="top-nav">
          {"".join(nav_links)}
        </nav>
        <form action="search.html" method="get" class="header-search-form">
          <input type="text" name="q" class="header-search-input" placeholder="Search sermons..." aria-label="Search sermons">
        </form>
      </div>
    </div>
  </header>
  """


# ==========================================
# INDIVIDUAL SERMON PAGES GENERATOR
# ==========================================
def generate_individual_sermon_pages(max_sermon=202):
    # Collect all valid existing sermon numbers (automatically handles 170 -> 183 gap)
    valid_sermons = []
    for i in range(1, max_sermon + 1):
        s_num = str(i).zfill(3)
        if find_sermon_file(s_num, "docx", SRC_DOCX_DIR) or find_sermon_file(s_num, "txt", SRC_TXT_DIR):
            valid_sermons.append(s_num)

    for idx, s_num in enumerate(valid_sermons):
        prev_num = valid_sermons[idx - 1] if idx > 0 else None
        next_num = valid_sermons[idx + 1] if idx < len(valid_sermons) - 1 else None

        docx_file = find_sermon_file(s_num, "docx", SRC_DOCX_DIR) or f"sermon_{s_num}_ENGLISH.docx"
        pdf_file = find_sermon_file(s_num, "pdf", SRC_PDF_DIR) or f"sermon_{s_num}.pdf"
        epub_file = find_sermon_file(s_num, "epub", SRC_EPUB_DIR) or f"sermon_{s_num}.epub"

        sermon_text = read_sermon_content(s_num)

        prev_link = f'<a href="sermon_{prev_num}.html" class="doc-btn">← Previous Sermon ({prev_num})</a>' if prev_num else '<span></span>'
        next_link = f'<a href="sermon_{next_num}.html" class="doc-btn">Next Sermon ({next_num}) →</a>' if next_num else '<span></span>'

        page_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Pahankanuwa Sermon {s_num} | Ven. Katukurunde Nanananda Thero</title>
  <style>{SITE_CSS}</style>
</head>
<body>

  {build_header_html('sermons')}

  <main class="homepage-content">

    <div class="sermon-nav-bar">
      {prev_link}
      <a href="sermons.html" class="doc-btn">📜 Back to All Sermons</a>
      {next_link}
    </div>

    <section class="hero" style="margin-bottom: 1rem;">
      <h1>Pahankanuwa Sermon {s_num}</h1>
      <p style="margin-top: 0.5rem;">
        <strong>Download Files:</strong>
        <a class="doc-btn" href="docx/{docx_file}" target="_blank">📄 DOCX</a>
        <a class="doc-btn" href="pdf/{pdf_file}" target="_blank">📕 PDF</a>
        <a class="doc-btn" href="epub/{epub_file}" target="_blank">📘 EPUB</a>
      </p>
    </section>

    <article class="sermon-reader">
      {sermon_text}
    </article>

    <div class="sermon-nav-bar" style="margin-top: 2rem;">
      {prev_link}
      <a href="sermons.html" class="doc-btn">📜 Back to All Sermons</a>
      {next_link}
    </div>

  </main>

  {FOOTER_HTML}

</body>
</html>
"""
        with open(DOCS_DIR / f"sermon_{s_num}.html", "w", encoding="utf-8") as f:
            f.write(page_html)

    print(f"Generated {len(valid_sermons)} individual sermon pages.")


# ==========================================
# PAGE GENERATORS
# ==========================================
def generate_index_page(metrics):
    index_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>English Translations of Pahankanuwa Sermons | Ven. Katukurunde Nanananda Thero</title>
  <style>{SITE_CSS}</style>
</head>
<body>

  {build_header_html('home')}

  <main class="homepage-content">

    <section class="hero">
      <h1>Pahankanuwa Sermon Series</h1>
      <h2>English Translation Collection (Sermons 001–202)</h2>
      <p style="text-align: center; font-weight: 500;">Original Sermons by Most Venerable Katukurunde Nanananda Thero</p>
      
      <div class="pali-quote">"Sabba-dānaṃ dhamma-dānaṃ jināti"<br>The Gift of Dhamma Excels All Gifts — Dhammapada 354</div>

      <div class="download-banner">
        <a href="{DOCX_COMPLETE}" class="btn" download>Download Complete DOCX</a>
        <a href="{PDF_COMPLETE}" class="btn" download>Download Complete PDF</a>
        <a href="{EPUB_COMPLETE}" class="btn" download>Download Complete EPUB</a>
        <a href="{ZIP_FILE}" class="btn btn-secondary" download>Download All (ZIP)</a>
      </div>

      <p>Welcome to the structured Dhamma Knowledge Base, cataloging and indexing the English translations of the Pahankanuwa Sermons delivered by Most Venerable Katukurunde Nanananda Thero. This portal provides direct cross-references between translated sermon passages, Dhamma subjects, and canonical Sutta citations.</p>

<div style="margin-top: 1.5rem; margin-bottom: 1.5rem; padding: 1.25rem; border-left: 5px solid #800000; background-color: #fff8f0; border-radius: 6px;">
    <h3 style="margin-bottom: 0.5rem; color: #800000;">
        Featured Collection: The Paticca Samuppada Series (Sermons 183–202)
    </h3>
    <p style="margin-bottom: 0;">
        Sermons 183–202 represent Ven. Katukurunde Nanananda Thero's profound exposition of Paṭicca Samuppāda (Dependent Arising), later published by the Venerable Thero as the book <em>The Law of Dependent Arising</em>.
    </p>
</div>

      <div style="margin-top: 1.5rem; margin-bottom: 1.5rem; padding: 1rem; border-left: 4px solid var(--header-bg); background-color: #fcfaf7; border-radius: 6px;">
        <h3 style="margin-bottom:0.5rem; color: var(--header-bg);">Sinhala Pahankanuwa Collection</h3>
        <p style="margin-bottom:0.75rem;">Original Sinhala editions and study resources.</p>
        <a href="https://dharmapress.github.io/" class="btn btn-secondary" target="_blank">Visit Dharmapress Sinhala Collection</a>
      </div>

      <p class="disclaimer">The translations presented in this knowledge base were prepared from digitised editions of the Pahankanuwa sermon series. Original sermon texts were derived from OCR-processed versions of publicly available PDF editions and translated through a collaborative workflow involving OCR correction, AI-assisted translation, and human review and proofreading. The project is intended as a freely accessible resource for Dhamma study and reference.<br><br><strong>Disclaimer:</strong> These translations are provided for educational and Dhamma-study purposes only and are not intended for commercial use. All credit for the original sermons belongs to Ven. Katukurunde Nanananda Thero and the original publishers. If you are a copyright holder and have concerns regarding the distribution of these materials, please contact the repository maintainer.</p>
    </section>

    <section class="stats-grid">
      <a href="sermons.html" class="stat-card">
        <span class="stat-number">{metrics['sermons_count']:,}</span>
        <span class="stat-label">Browse All Sermons</span>
      </a>
      <a href="downloads.html" class="stat-card">
        <span class="stat-number">4 Formats</span>
        <span class="stat-label">Downloads Page</span>
      </a>
      <a href="statistics.html" class="stat-card">
        <span class="stat-number">100%</span>
        <span class="stat-label">View Collection Stats</span>
      </a>
      <a href="search.html" class="stat-card">
        <span class="stat-number">🔍</span>
        <span class="stat-label">Search Entire Collection</span>
      </a>
    </section>

  </main>

  {FOOTER_HTML}

</body>
</html>
"""
    with open(DOCS_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(index_html)


def generate_downloads_page(metrics):
    individual_items = []
    for i in range(1, 203):
        pad_num = str(i).zfill(3)
        if not (find_sermon_file(pad_num, "docx", SRC_DOCX_DIR) or find_sermon_file(pad_num, "txt", SRC_TXT_DIR)):
            continue
            
        docx_file = find_sermon_file(pad_num, "docx", SRC_DOCX_DIR) or f"sermon_{pad_num}_ENGLISH.docx"
        pdf_file = find_sermon_file(pad_num, "pdf", SRC_PDF_DIR) or f"sermon_{pad_num}.pdf"
        epub_file = find_sermon_file(pad_num, "epub", SRC_EPUB_DIR) or f"sermon_{pad_num}.epub"

        item_html = f"""
        <div class="download-item">
            <div class="download-info">
                <h4>Pahankanuwa Sermon #{i}</h4>
                <p>Individual sermon document files</p>
            </div>
            <div class="download-links">
                <a href="docx/{docx_file}" class="btn btn-secondary" download>DOCX</a>
                <a href="pdf/{pdf_file}" class="btn btn-secondary" download>PDF</a>
                <a href="epub/{epub_file}" class="btn btn-secondary" download>EPUB</a>
            </div>
        </div>
        """
        individual_items.append(item_html)

    downloads_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Downloads & Documents | Ven. Katukurunde Nanananda Thero</title>
  <style>{SITE_CSS}</style>
</head>
<body>

  {build_header_html('downloads')}

  <div class="search-container">
    <input type="text" id="searchInput" class="search-input" onkeyup="filterData()" placeholder="Search downloads by sermon number or format...">
  </div>

  <main class="homepage-content">
    <section class="hero">
      <h1>Downloads & Documents</h1>
      <p>Access full manuscript collections or individual sermon files in your preferred format.</p>
    </section>

    <div class="download-list">
      <div class="download-item" style="border-left: 4px solid var(--header-bg);">
        <div class="download-info">
          <h4>Complete Sermon Collection (All Sermons)</h4>
          <p>Complete set available in DOCX, PDF, EPUB, and a combined ZIP archive.</p>
        </div>
        <div class="download-links">
          <a href="{ZIP_FILE}" class="btn" download>ZIP Archive</a>
          <a href="{DOCX_COMPLETE}" class="btn btn-secondary" download>DOCX</a>
          <a href="{PDF_COMPLETE}" class="btn btn-secondary" download>PDF</a>
          <a href="{EPUB_COMPLETE}" class="btn btn-secondary" download>EPUB</a>
        </div>
      </div>

      <h2 style="font-family:'Georgia', serif; color: var(--primary-color); margin-top: 2rem;">Individual Sermon Downloads</h2>
      {"".join(individual_items)}
    </div>
  </main>

  {FOOTER_HTML}
  {SEARCH_JS}

</body>
</html>
"""
    with open(DOCS_DIR / "downloads.html", "w", encoding="utf-8") as f:
        f.write(downloads_html)


def generate_audio_page():
    audio_files = set()
    for audio_dir in [SRC_AUDIO_DIR, DOCS_AUDIO_DIR]:
        if audio_dir.exists():
            for f in audio_dir.iterdir():
                if f.is_file() and (f.suffix.lower() == '.mp3' or 'sermon' in f.name.lower() or 'pahankanuwa' in f.name.lower() or 'paticca' in f.name.lower()):
                    dest_name = f.name if f.name.lower().endswith(".mp3") else f"{f.name}.mp3"
                    audio_files.add(dest_name)

    paticca_files = []
    pahankanuwa_files = []

    for file_name in audio_files:
        if "paticca" in file_name.lower():
            paticca_files.append(file_name)
        else:
            pahankanuwa_files.append(file_name)

    def sort_key(filename):
        numbers = re.findall(r'\d+', filename)
        return int(numbers[0]) if numbers else 0

    paticca_sorted = sorted(paticca_files, key=sort_key)
    pahankanuwa_sorted = sorted(pahankanuwa_files, key=sort_key)

    def build_playlist_html(file_list, title_prefix, description_prefix):
        if not file_list:
            return "<p style='color: var(--text-muted); font-style: italic;'>No recordings currently available in this series.</p>"
        
        cards = []
        for audio_file in file_list:
            match = re.search(r'(\d{1,3})', audio_file)
            item_num = match.group(1).zfill(3) if match else "---"
            
            card_html = f"""
<div class="download-item">
  <div class="download-info">
    <h4>🎧 {title_prefix} #{item_num}</h4>
    <p>{description_prefix} ({audio_file})</p>
  </div>
  <div class="download-links">
    <a href="audio/{audio_file}" class="btn btn-secondary" download>Download MP3</a>
  </div>
</div>
<audio controls style="width:100%; margin-top: 0.5rem; margin-bottom: 1.5rem;">
  <source src="audio/{audio_file}" type="audio/mpeg">
  Your browser does not support the audio element.
</audio>
"""
            cards.append(card_html)
        return "\n".join(cards)

    paticca_html = build_playlist_html(paticca_sorted, "Paticca Samuppada", "Paticca Samuppada Audio Series")
    pahankanuwa_html = build_playlist_html(pahankanuwa_sorted, "Pahankanuwa Sermon", "English Audio Narration")

    page_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Audio Recordings | Ven. Katukurunde Nanananda Thero</title>
<style>{SITE_CSS}</style>
</head>
<body>

{build_header_html('audio')}

<main class="homepage-content">
<section class="hero">

<h1>Audio Recordings</h1>
<p>Listen to English audio recordings of Venerable Katukurunde Nanananda Thero's Dhamma teachings.</p>

<h2 style="color: var(--primary-color); margin-top: 1.5rem; border-bottom: 2px solid var(--border-color); padding-bottom: 0.5rem;">
  Paticca Samuppada Series
</h2>
{paticca_html}

<h2 style="color: var(--primary-color); margin-top: 2.5rem; border-bottom: 2px solid var(--border-color); padding-bottom: 0.5rem;">
  Pahankanuwa Sermons
</h2>
{pahankanuwa_html}

</section>
</main>

{FOOTER_HTML}

</body>
</html>
"""
    with open(DOCS_DIR / "audio.html", "w", encoding="utf-8") as f:
        f.write(page_html)


def generate_data_subpage(filename, tab_key, page_title, description, df_data):
    if df_data is not None and not df_data.empty:
        df_clean = df_data.fillna("").copy()
        headers_html = "".join([f"<th>{col}</th>" for col in df_clean.columns])
        
        rows_list = []
        for _, row in df_clean.iterrows():
            cells = "".join([f"<td>{row[col]}</td>" for col in df_clean.columns])
            rows_list.append(f"<tr>{cells}</tr>")
        
        rows_html = "".join(rows_list)
        table_html = f"""
        <div class="table-container">
            <table class="data-table">
                <thead><tr>{headers_html}</tr></thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        """
    else:
        table_html = "<div class='hero'><p>No records found.</p></div>"

    page_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{page_title} | Ven. Katukurunde Nanananda Thero</title>
  <style>{SITE_CSS}</style>
</head>
<body>

  {build_header_html(tab_key)}

  <div class="search-container">
    <input type="text" id="searchInput" class="search-input" onkeyup="filterData()" placeholder="Search this index...">
  </div>

  <main class="homepage-content">
    <section class="hero">
      <h1>{page_title}</h1>
      <p>{description}</p>
    </section>

    {table_html}
  </main>

  {FOOTER_HTML}
  {SEARCH_JS}

</body>
</html>
"""
    with open(DOCS_DIR / filename, "w", encoding="utf-8") as f:
        f.write(page_html)


def generate_statistics_page(metrics, datasets):
    sermons_df = datasets.get("sermons", pd.DataFrame())
    sermon_rows_list = []
    if not sermons_df.empty:
        for idx, row in sermons_df.iterrows():
            sermon_title = row.get('Sermon', f'Sermon {idx+1}')
            excerpt = row.get('Opening Passage Excerpt', '')
            source = row.get('Likely Source', '')
            doc_link = row.get('Document Link', '')
            
            sermon_rows_list.append(f"""
            <tr>
              <td>{sermon_title}</td>
              <td>{source}</td>
              <td>{excerpt}</td>
              <td>{doc_link}</td>
            </tr>
            """)
    
    sermon_table_body = "".join(sermon_rows_list)

    page_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Collection Statistics | Ven. Katukurunde Nanananda Thero</title>
  <style>{SITE_CSS}</style>
</head>
<body>

  {build_header_html('statistics')}

  <div class="search-container">
    <input type="text" id="searchInput" class="search-input" onkeyup="filterData()" placeholder="Search statistics or jump to sermon...">
  </div>

  <main class="homepage-content">
    <section class="hero">
      <h1>Collection Statistics & Direct Sermon Access</h1>
      <p>Overview of the Pahankanuwa translation collection with direct links to sermon documents at the end of the tree.</p>
    </section>

    <section class="stats-grid">
      <a href="#sermons-tree" class="stat-card">
        <span class="stat-number">{metrics['sermons_count']:,}</span>
        <span class="stat-label">Sermons (Click to Access Documents)</span>
      </a>
      <a href="#sermons-tree" class="stat-card">
        <span class="stat-number">100%</span>
        <span class="stat-label">Document Availability</span>
      </a>
    </section>

    <section id="sermons-tree">
      <h2 style="font-family:'Georgia', serif; color: var(--primary-color); margin-bottom: 1rem;">Sermons Index & Document Links</h2>
      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <th>Sermon</th>
              <th>Likely Canonical Source</th>
              <th>Opening Passage Excerpt</th>
              <th>Document Link</th>
            </tr>
          </thead>
          <tbody>
            {sermon_table_body}
          </tbody>
        </table>
      </div>
    </section>

  </main>

  {FOOTER_HTML}
  {SEARCH_JS}

</body>
</html>
"""
    with open(DOCS_DIR / "statistics.html", "w", encoding="utf-8") as f:
        f.write(page_html)

def generate_search_index():
    search_list = []
    for i in range(1, 203):
        s_num = str(i).zfill(3)
        # Check if the sermon file exists before indexing (handles 170 -> 183 gap)
        docx_file = find_sermon_file(s_num, "docx", SRC_DOCX_DIR)
        txt_file = find_sermon_file(s_num, "txt", SRC_TXT_DIR)
        if not (docx_file or txt_file):
            continue

        title = f"Pahankanuwa Sermon {s_num}"
        content = read_sermon_content(s_num)
        clean_content = re.sub(r'<[^>]+>', '', content)
        
        search_item = {
            "sermon": s_num,
            "title": title,
            "content": clean_content,
            "normalized": normalize_text(title + " " + clean_content)
        }
        search_list.append(search_item)
        
    with open(DOCS_DIR / "search_index.json", "w", encoding="utf-8") as f:
        json.dump(search_list, f, ensure_ascii=False, indent=2)
    print("Generated search_index.json successfully.")

def generate_search_page():
    search_html = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Search | Ven. Katukurunde Nanananda Thero</title>

<style>
__SITE_CSS__
</style>
</head>

<body>

__HEADER_HTML__

<main class="homepage-content">

<section class="hero">

<h1>Search the Sermon Collection</h1>

<p>
Search all translated sermons,
including the Paticca Samuppada series.
Searches ignore diacritics,
so typing "paticca" will find "Paṭicca".
</p>

<input
id="searchBox"
class="search-input"
placeholder="Search..."
>

<div id="results"></div>

</section>

</main>

<script>

let searchData=[];

function normalize(str){
    return str
    .normalize("NFD")
    .replace(/[\\u0300-\\u036f]/g,"")
    .toLowerCase();
}

function getSnippet(content, query, maxLength = 250) {
    const lowerContent = content.toLowerCase();
    const index = lowerContent.indexOf(query);
    if (index === -1) {
        return content.substring(0, maxLength);
    }
    const start = Math.max(0, index - 80);
    const end = Math.min(content.length, start + maxLength);
    let snippet = content.substring(start, end);
    if (start > 0) snippet = "..." + snippet;
    if (end < content.length) snippet = snippet + "...";
    return snippet;
}

function performSearch() {
    const box = document.getElementById('searchBox');
    const query = normalize(box.value.trim());
    const results = document.getElementById('results');

    if (query.length < 2) {
        results.innerHTML = "";
        return;
    }

    const matches = searchData.filter(x => x.normalized.includes(query)).slice(0, 50);

    if (matches.length === 0) {
        results.innerHTML = "<p style='margin-top: 1.5rem; color: var(--text-muted);'>No matching sermons found.</p>";
        return;
    }

    results.innerHTML = matches.map(m => `
<div class="download-item">
  <div class="download-info">
    <h4>
      <a href="sermon_${m.sermon}.html">
        ${m.title}
      </a>
    </h4>
    <p>
      ${getSnippet(m.content, query, 250)}
    </p>
  </div>
</div>
`).join("");
}

fetch('search_index.json')
.then(r => r.json())
.then(data => {
    searchData = data;
    
    // Check URL parameters (e.g. ?q=paticca)
    const urlParams = new URLSearchParams(window.location.search);
    const initialQuery = urlParams.get('q') || urlParams.get('search') || "";
    if (initialQuery) {
        const box = document.getElementById('searchBox');
        box.value = initialQuery;
        performSearch();
    }
});

document.getElementById('searchBox').addEventListener('input', performSearch);

</script>

</body>

</html>
""".replace("__SITE_CSS__", SITE_CSS).replace("__HEADER_HTML__", build_header_html('search'))

    with open(
        DOCS_DIR / "search.html",
        "w",
        encoding="utf-8"
    ) as f:
        f.write(search_html)


# ==========================================
# SITEMAP GENERATOR
# ==========================================
def generate_sitemap(max_sermon=202, base_url="https://chamaniw.github.io/Pahankanuwa-english-translations/"):
    """
    Generates a sitemap.xml file inside the docs folder listing all main pages and individual sermon URLs.
    Update base_url if your custom domain or GitHub Pages URL differs.
    """
    pages = [
        "",
        "index.html",
        "sermons.html",
        "downloads.html",
        "audio.html",
        "search.html",
        "statistics.html"
    ]
    
    # Collect all valid individual sermon pages (automatically handles gaps like 170 -> 183)
    for i in range(1, max_sermon + 1):
        s_num = str(i).zfill(3)
        if find_sermon_file(s_num, "docx", SRC_DOCX_DIR) or find_sermon_file(s_num, "txt", SRC_TXT_DIR):
            pages.append(f"sermon_{s_num}.html")
            
    sitemap_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
    ]
    
    for page in pages:
        loc = f"{base_url.rstrip('/')}/{page}" if page else f"{base_url.rstrip('/')}/"
        sitemap_lines.append('  <url>')
        sitemap_lines.append(f'    <loc>{loc}</loc>')
        sitemap_lines.append('  </url>')
        
    sitemap_lines.append('</urlset>')
    
    sitemap_path = DOCS_DIR / "sitemap.xml"
    with open(sitemap_path, "w", encoding="utf-8") as f:
        f.write("\n".join(sitemap_lines))
    print(f"Generated sitemap.xml with {len(pages)} URLs successfully.")


# ==========================================
# FILE COPY & ASSET MANAGEMENT
# ==========================================
def copy_static_assets():
    if SRC_DOCX_DIR.exists():
        docx_files = list(SRC_DOCX_DIR.glob("*.docx"))
        for df in docx_files:
            shutil.copy2(df, DOCS_DOCX_DIR / df.name)
        print(f"Copied {len(docx_files)} DOCX files into docs/docx/")

    if SRC_PDF_DIR.exists():
        pdf_files = list(SRC_PDF_DIR.glob("*.pdf"))
        for pf in pdf_files:
            shutil.copy2(pf, DOCS_PDF_DIR / pf.name)
        print(f"Copied {len(pdf_files)} PDF files into docs/pdf/")

    if SRC_EPUB_DIR.exists():
        epub_files = list(SRC_EPUB_DIR.glob("*.epub"))
        for ef in epub_files:
            shutil.copy2(ef, DOCS_EPUB_DIR / ef.name)
        print(f"Copied {len(epub_files)} EPUB files into docs/epub/")

    if SRC_AUDIO_DIR.exists():
        audio_files = list(SRC_AUDIO_DIR.glob("*"))
        copied_count = 0
        for af in audio_files:
            if af.is_file():
                dest_name = af.name if af.name.lower().endswith(".mp3") else f"{af.name}.mp3"
                shutil.copy2(af, DOCS_AUDIO_DIR / dest_name)
                copied_count += 1
        print(f"Copied {copied_count} Audio files into docs/audio/")

    for file_name in [ZIP_FILE, DOCX_COMPLETE, PDF_COMPLETE, EPUB_COMPLETE]:
        src_file = BASE_DIR / file_name
        if src_file.exists():
            shutil.copy2(src_file, DOCS_DIR / file_name)


# ==========================================
# MAIN EXECUTION
# ==========================================
if __name__ == "__main__":
    print("==========================================")
    print(" Building Pahankanuwa Website")
    print("==========================================\n")
    
    metrics, datasets = load_and_process_excel_data()
    
    generate_index_page(metrics)
    generate_data_subpage(
        "sermons.html", "sermons", "Browse Sermons",
        "Catalog of Pahankanuwa sermon translations with opening excerpts, canonical sources, and document links.",
        datasets["sermons"]
    )
    generate_downloads_page(metrics)
    generate_audio_page()
    generate_statistics_page(metrics, datasets)
    
    # Generate individual sermon pages and handle gaps automatically
    generate_individual_sermon_pages(202)

    # Generate search index JSON file required by search.html (with gap skipping)
    generate_search_index()

    generate_search_page()

    # Generate sitemap.xml for Google Search Console
    generate_sitemap(202)

    copy_static_assets()
    
    print("\n==========================================")
    print(" Generation Complete! Open C:\\Pahankanuwa\\docs\\index.html in browser.")
    print("==========================================")