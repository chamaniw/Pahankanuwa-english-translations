import os
import shutil
import re
from pathlib import Path
import pandas as pd

# ==========================================
# PATH & DIRECTORY CONFIGURATION
# ==========================================
BASE_DIR = Path(r"C:\Pahankanuwa")
DOCS_DIR = BASE_DIR / "docs"
DOCX_DIR = BASE_DIR / "docx"
DOCS_DOCX_DIR = DOCS_DIR / "docx"
ZIP_FILE = BASE_DIR / "Pahankanuwa_English_Translations.zip"

DOCS_DIR.mkdir(parents=True, exist_ok=True)
DOCS_DOCX_DIR.mkdir(parents=True, exist_ok=True)


# ==========================================
# HELPER FUNCTIONS
# ==========================================
def parse_and_link_sermons(val):
    """
    Converts sermon numbers or string IDs into interactive DOCX link buttons.
    """
    if pd.isna(val) or not str(val).strip() or str(val).lower() == "nan":
        return ""
    
    val_str = str(val).strip()
    
    # Single numeric sermon string (e.g., "1", "001", "150")
    if val_str.isdigit() and len(val_str) <= 3:
        s_num = val_str.zfill(3)
        doc_filename = f"sermon_{s_num}_ENGLISH.docx"
        return f'<a class="doc-btn" href="docx/{doc_filename}" target="_blank">📄 Sermon {s_num}</a>'

    # Multi-sermon references separated by comma or semicolon
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


# ==========================================
# DATA LOADING & PROCESSOR
# ==========================================
def load_and_process_excel_data():
    scratch_file = BASE_DIR / "pahankanuwa_e_from_scratch_4.xlsx"
    idx_file = BASE_DIR / "Nanananda_Dhamma_Index_V3.xlsx"
    kb_file = BASE_DIR / "Nanananda_Knowledge_Base_V1.xlsx"

    datasets = {
        "sermons": pd.DataFrame(),
        "citations": pd.DataFrame(),
        "statistics": pd.DataFrame()
    }

    # 1. Load Sermons (Para_7 for Opening Excerpt, Para_9 for Likely Source)
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
            docx_target = f"sermon_{s_num}_ENGLISH.docx"

            # Para_7 for Opening Passage Excerpt
            para_7 = str(row['Para_7']).strip() if 'Para_7' in row and pd.notna(row['Para_7']) else ""
            excerpt = para_7 if para_7 and para_7.lower() != 'nan' else "No excerpt available"

            # Para_9 for Likely Source
            para_9 = str(row['Para_9']).strip() if 'Para_9' in row and pd.notna(row['Para_9']) else ""
            likely_source = para_9 if para_9 and para_9.lower() != 'nan' else "Unspecified Canonical Source"

            sermons_list.append({
                "Sermon": f"<strong>{clean_title}</strong>",
                "Opening Passage Excerpt": excerpt.replace('\n', '<br>'),
                "Likely Source": likely_source,
                "Document Link": f'<a class="doc-btn" href="docx/{docx_target}" target="_blank">📄 Open DOCX Document</a>'
            })
        datasets["sermons"] = pd.DataFrame(sermons_list)

    # 2. Load Citations
    if idx_file.exists():
        xls_idx = pd.ExcelFile(idx_file)
        if 'Pali_Citations' in xls_idx.sheet_names:
            df_cit = pd.read_excel(xls_idx, sheet_name='Pali_Citations')
            df_cit['Sermon Link'] = df_cit['Sermon'].apply(parse_and_link_sermons)
            datasets["citations"] = df_cit[['Citation', 'Sermon Link', 'Type', 'Likely Source']].drop_duplicates()

    # 3. Load Statistics
    if kb_file.exists():
        xls_kb = pd.ExcelFile(kb_file)
        if 'Statistics' in xls_kb.sheet_names:
            datasets["statistics"] = pd.read_excel(xls_kb, sheet_name='Statistics')

    metrics = {
        "sermons_count": len(datasets["sermons"]) if not datasets["sermons"].empty else 170,
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

.top-nav {
    display: flex;
    gap: 0.5rem;
    flex-wrap: wrap;
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
    margin: 1.5rem auto 3rem auto;
    padding: 0 1.5rem;
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
.hero p { margin-bottom: 1rem; color: #333; font-size: 0.98rem; text-align: justify; }
.hero p.disclaimer { font-size: 0.88rem; color: var(--text-muted); font-style: italic; border-top: 1px solid var(--border-color); padding-top: 1rem; margin-top: 1.5rem; }

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

footer {
    text-align: center;
    padding: 2rem;
    font-size: 0.85rem;
    color: var(--text-muted);
    border-top: 1px solid var(--border-color);
    margin-top: 3rem;
}
.zip-btn {
    display: inline-block;
    background-color: var(--header-bg);
    color: #ffffff !important;
    padding: 0.6rem 1.2rem;
    border-radius: 6px;
    font-weight: bold;
    text-decoration: none;
    font-size: 0.95rem;
    transition: background-color 0.2s;
}
.zip-btn:hover {
    background-color: #5c0000;
}
"""

SEARCH_JS = """
<script>
function filterData() {
    var input = document.getElementById('searchInput');
    var filter = input.value.toLowerCase();
    
    var rows = document.querySelectorAll('.data-table tbody tr');
    rows.forEach(function(row) {
        var text = row.innerText.toLowerCase();
        row.style.display = text.includes(filter) ? '' : 'none';
    });
}
</script>
"""

FOOTER_HTML = """
  <footer>
    <p style="margin-bottom: 1rem;">
      <a class="zip-btn" href="Pahankanuwa_English_Translations.zip" download>
        📦 Download All Sermons (.ZIP Archive)
      </a>
    </p>
    <p>Ven. Katukurunde Nanananda Thero Dhamma Knowledge Base &copy; 2026.</p>
  </footer>
"""


# ==========================================
# HEADER TEMPLATE (CONCEPTS & SUTTAS REMOVED)
# ==========================================
def build_header_html(active_tab="home"):
    tabs = [
        ("home", "Home", "index.html"),
        ("sermons", "Sermons", "sermons.html"),
        ("citations", "Citations", "citations.html"),
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
      <nav class="top-nav">
        {"".join(nav_links)}
      </nav>
    </div>
  </header>
  """


# ==========================================
# HOMEPAGE GENERATOR
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

  <div class="search-container">
    <input type="text" id="searchInput" class="search-input" onkeyup="filterData()" placeholder="Search sermon index by title, citation, or source...">
  </div>

  <main class="homepage-content">

    <section class="hero">
      <h1>English Translations of Pahankanuwa Sermons</h1>
      <h2>Venerable Katukurunde Nanananda Thero</h2>
      
      <p>Welcome to the structured Dhamma Knowledge Base, cataloging and indexing the English translations of the Pahankanuwa Sermons delivered by Most Venerable Katukurunde Nanananda Thero. This portal provides direct cross-references between translated sermon passages, Dhamma subjects, and canonical Sutta citations.</p>
      
      <p>The translations presented in this knowledge base were prepared from digitised editions of the Pahankanuwa sermon series. Original sermon texts were derived from OCR-processed versions of publicly available PDF editions and translated through a collaborative workflow involving OCR correction, AI-assisted translation, and human review and proofreading. The project is intended as a freely accessible resource for Dhamma study and reference.</p>
      
      <p class="disclaimer"><strong>Disclaimer:</strong> These translations are provided for educational and Dhamma-study purposes only and are not intended for commercial use. All credit for the original sermons belongs to Ven. Katukurunde Nanananda Thero and the original publishers. If you are a copyright holder and have concerns regarding the distribution of these materials, please contact the repository maintainer.</p>
    </section>

    <section class="stats-grid">
      <a href="sermons.html" class="stat-card">
        <span class="stat-number">{metrics['sermons_count']:,}</span>
        <span class="stat-label">Browse All Sermons</span>
      </a>
      <a href="citations.html" class="stat-card">
        <span class="stat-number">{metrics['citations_count']:,}</span>
        <span class="stat-label">Browse Citations Index</span>
      </a>
      <a href="statistics.html" class="stat-card">
        <span class="stat-number">100%</span>
        <span class="stat-label">View Full Collection Stats</span>
      </a>
    </section>

  </main>

  {FOOTER_HTML}

  {SEARCH_JS}

</body>
</html>
"""
    with open(DOCS_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(index_html)
    print("Generated: docs/index.html")


# ==========================================
# SUBPAGE GENERATOR
# ==========================================
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
    print(f"Generated page: docs/{filename}")


# ==========================================
# CLICKABLE STATISTICS PAGE GENERATOR
# ==========================================
def generate_statistics_page(metrics, datasets):
    """
    Renders Statistics with direct sermon document links down the tree.
    """
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
      <a href="citations.html" class="stat-card">
        <span class="stat-number">{metrics['citations_count']:,}</span>
        <span class="stat-label">Citations (Click to View Index)</span>
      </a>
      <a href="#sermons-tree" class="stat-card">
        <span class="stat-number">100%</span>
        <span class="stat-label">DOCX Availability</span>
      </a>
    </section>

    <section id="sermons-tree">
      <h2 style="font-family:'Georgia', serif; color: var(--primary-color); margin-bottom: 1rem;">Sermons Index & DOCX Document Links</h2>
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
    print("Generated clickable: docs/statistics.html")


# ==========================================
# FILE COPY & ASSET MANAGEMENT
# ==========================================
def copy_static_assets():
    if DOCX_DIR.exists():
        docx_files = list(DOCX_DIR.glob("*.docx"))
        for df in docx_files:
            shutil.copy2(df, DOCS_DOCX_DIR / df.name)
        print(f"Copied {len(docx_files)} DOCX files into docs/docx/")

    if ZIP_FILE.exists():
        shutil.copy2(ZIP_FILE, DOCS_DIR / ZIP_FILE.name)
        print("Copied ZIP archive to docs/")


# ==========================================
# MAIN EXECUTION
# ==========================================
if __name__ == "__main__":
    print("==========================================")
    print(" Building Pahankanuwa Website")
    print("==========================================\n")
    
    # 1. Load Data
    metrics, datasets = load_and_process_excel_data()
    
    # 2. Build Home Page
    generate_index_page(metrics)
    
    # 3. Build Sermons Page
    generate_data_subpage(
        "sermons.html", "sermons", "Browse Sermons",
        "Catalog of Pahankanuwa sermon translations with opening excerpts (Para_7), canonical sources (Para_9), and DOCX links.",
        datasets["sermons"]
    )
    
    # 4. Build Citations Page
    generate_data_subpage(
        "citations.html", "citations", "Browse Citations Index",
        "Opening passages and textual citations with direct links to corresponding sermon translations.",
        datasets["citations"]
    )
    
    # 5. Build Clickable Statistics Page
    generate_statistics_page(metrics, datasets)
    
    # 6. Copy Static Assets
    copy_static_assets()
    
    print("\n==========================================")
    print(" Generation Complete! Open C:\\Pahankanuwa\\docs\\index.html in browser.")
    print("==========================================")