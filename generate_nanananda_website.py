import os
import re
import shutil
import json
import pandas as pd
from pathlib import Path

# ==============================================================================
# CONFIGURATION & FILE PATHS
# ==============================================================================
INPUT_PATH = r"C:\Users\chama\OneDrive\Pahankanuwa Project\Editing English Translations\Nanananda_Knowledge_Base_V1.xlsx"
OUTPUT_DIR = Path("docs")
GITHUB_DOCX_BASE_URL = "https://github.com/chamaniw/Pahankanuwa-english-translations/blob/main/docx/"

# Ensure required output directory tree exists
os.makedirs(OUTPUT_DIR / "sermons", exist_ok=True)
os.makedirs(OUTPUT_DIR / "subjects", exist_ok=True)
os.makedirs(OUTPUT_DIR / "suttas", exist_ok=True)
os.makedirs(OUTPUT_DIR / "citations", exist_ok=True)


# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================
def slugify(text: str) -> str:
    """Converts strings/subjects/suttas into clean, URL-safe filenames."""
    if not isinstance(text, str) or not text.strip():
        return "unknown"
    clean = text.strip().lower()
    clean = re.sub(r'[\s/\\\-\:\,\.\(\)]+', '_', clean)
    clean = re.sub(r'[^a-z0-9_]', '', clean)
    clean = re.sub(r'_+', '_', clean).strip('_')
    return clean or "item"

def format_sermon_filename(sermon_id: str) -> str:
    """Formats raw Sermon IDs (e.g., '1', 'Sermon 001') into 'sermon_001_ENGLISH.html'."""
    numbers = re.findall(r'\d+', str(sermon_id))
    if numbers:
        num = int(numbers[0])
        return f"sermon_{num:03d}_ENGLISH.html"
    return f"sermon_{slugify(str(sermon_id))}.html"

def format_sermon_docx_name(sermon_id: str) -> str:
    """Formats raw Sermon IDs into standard DOCX file name matching GitHub repository structure."""
    numbers = re.findall(r'\d+', str(sermon_id))
    if numbers:
        num = int(numbers[0])
        return f"sermon_{num:03d}_ENGLISH"
    return slugify(str(sermon_id))

def get_base_html(title: str, content: str, rel_path_to_root: str = "") -> str:
    """Generates consistent HTML page structure with navbar, search box, and footer."""
    site_title_text = "Pahankanuwa Sermons (English Translations) | Ven. Katukurunde Nanananda Thero"
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} | {site_title_text}</title>
    <link rel="stylesheet" href="{rel_path_to_root}style.css">
    <script defer src="{rel_path_to_root}search.js"></script>
</head>
<body>
    <header class="site-header">
        <div class="container header-container">
            <div class="site-branding">
                <h1 class="site-title"><a href="{rel_path_to_root}index.html">Ven. Katukurunde Nanananda Thero</a></h1>
                <span class="site-tagline">English Translations of Pahankanuwa Sermons</span>
            </div>
            <nav class="main-nav">
                <a href="{rel_path_to_root}index.html">Home</a>
                <a href="{rel_path_to_root}sermons/index.html">Sermons</a>
                <a href="{rel_path_to_root}subjects/index.html">Subjects</a>
                <a href="{rel_path_to_root}suttas/index.html">Suttas</a>
                <a href="{rel_path_to_root}citations/index.html">Citations</a>
                <a href="{rel_path_to_root}statistics.html">Statistics</a>
            </nav>
        </div>
    </header>

    <div class="search-container container">
        <input type="text" id="searchInput" placeholder="Search English translations by sermon title, subject, sutta, or citation..." onkeyup="performSearch()">
        <div id="searchResults" class="search-results"></div>
    </div>

    <main class="container main-content">
        {content}
    </main>

    <footer class="site-footer">
        <div class="container">
            <p>English Translations of Pahankanuwa Sermons by Venerable Katukurunde Nanananda Thero &copy; Knowledge Base Static Generator</p>
        </div>
    </footer>
</body>
</html>
"""

# ==============================================================================
# CSS & JS ASSET GENERATION
# ==============================================================================
def write_assets():
    """Generates external CSS stylesheet and client-side JS search engine."""
    css_content = """/* Academic & Modern Styling for Dhamma Knowledge Base */
:root {
    --primary: #8b0000;
    --primary-hover: #a00000;
    --bg-main: #fcfbf7;
    --card-bg: #ffffff;
    --text-dark: #2c2c2c;
    --text-muted: #666666;
    --border-color: #e2e0d8;
    --accent: #d4af37;
}

body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Georgia, serif;
    background-color: var(--bg-main);
    color: var(--text-dark);
    line-height: 1.6;
    margin: 0;
    padding: 0;
}

.container {
    max-width: 1050px;
    margin: 0 auto;
    padding: 0 20px;
}

.site-header {
    background-color: var(--primary);
    color: white;
    padding: 1.2rem 0;
    box-shadow: 0 2px 5px rgba(0,0,0,0.1);
}

.header-container {
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
}

.site-branding {
    display: flex;
    flex-direction: column;
}

.site-title {
    margin: 0;
    font-size: 1.4rem;
    font-weight: 600;
}

.site-title a {
    color: white;
    text-decoration: none;
}

.site-tagline {
    font-size: 0.85rem;
    color: #f0e6d2;
    font-style: italic;
    margin-top: 2px;
}

.main-nav a {
    color: #f0e6d2;
    text-decoration: none;
    margin-left: 20px;
    font-weight: 500;
    transition: color 0.2s;
}

.main-nav a:hover {
    color: white;
    text-decoration: underline;
}

.search-container {
    margin-top: 25px;
    position: relative;
}

#searchInput {
    width: 100%;
    padding: 12px 16px;
    font-size: 1rem;
    border: 1px solid var(--border-color);
    border-radius: 6px;
    box-shadow: inset 0 1px 3px rgba(0,0,0,0.05);
    box-sizing: border-box;
}

.search-results {
    position: absolute;
    top: 100%;
    left: 20px;
    right: 20px;
    background: white;
    border: 1px solid var(--border-color);
    border-radius: 0 0 6px 6px;
    max-height: 400px;
    overflow-y: auto;
    z-index: 1000;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
}

.search-result-item {
    padding: 12px;
    border-bottom: 1px solid #eee;
}

.search-result-item:last-child {
    border-bottom: none;
}

.search-result-item a {
    color: var(--primary);
    text-decoration: none;
    font-weight: bold;
}

.main-content {
    padding: 30px 20px 60px 20px;
}

.card {
    background: var(--card-bg);
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 24px;
    margin-bottom: 24px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}

.grid-4 {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 20px;
    margin-bottom: 30px;
}

.stat-card {
    background: white;
    border: 1px solid var(--border-color);
    border-top: 4px solid var(--primary);
    padding: 20px;
    text-align: center;
    border-radius: 6px;
}

.stat-number {
    font-size: 2.2rem;
    font-weight: bold;
    color: var(--primary);
}

.stat-label {
    color: var(--text-muted);
    font-size: 0.95rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 15px;
    background: white;
}

th, td {
    text-align: left;
    padding: 12px;
    border-bottom: 1px solid var(--border-color);
}

th {
    background-color: #f4f2eb;
    color: var(--text-dark);
}

a {
    color: var(--primary);
    text-decoration: none;
}

a:hover {
    text-decoration: underline;
}

.btn-docx {
    display: inline-block;
    background-color: #2b579a;
    color: #ffffff;
    padding: 10px 18px;
    border-radius: 5px;
    font-weight: 600;
    text-decoration: none;
    margin-top: 8px;
    transition: background-color 0.2s ease;
}

.btn-docx:hover {
    background-color: #1e3d6b;
    color: #ffffff;
    text-decoration: none;
}

.back-link {
    display: inline-block;
    margin-bottom: 20px;
    font-weight: 500;
}

.site-footer {
    border-top: 1px solid var(--border-color);
    padding: 20px 0;
    text-align: center;
    color: var(--text-muted);
    font-size: 0.9rem;
    background: #f4f2eb;
}
"""
    with open(OUTPUT_DIR / "style.css", "w", encoding="utf-8") as f:
        f.write(css_content)

    js_content = """let searchData = [];

fetch('/search_index.json')
    .then(response => response.json())
    .then(data => { searchData = data; })
    .catch(() => {
        fetch('search_index.json')
            .then(res => res.json())
            .then(data => { searchData = data; });
    });

function performSearch() {
    const input = document.getElementById('searchInput').value.toLowerCase().trim();
    const resultsContainer = document.getElementById('searchResults');
    resultsContainer.innerHTML = '';

    if (input.length < 2) {
        return;
    }

    const matches = searchData.filter(item => 
        item.title.toLowerCase().includes(input) || 
        item.text.toLowerCase().includes(input)
    ).slice(0, 15);

    if (matches.length === 0) {
        resultsContainer.innerHTML = '<div class="search-result-item">No results found.</div>';
        return;
    }

    matches.forEach(item => {
        const div = document.createElement('div');
        div.className = 'search-result-item';
        div.innerHTML = `<a href="${item.url}">${item.title}</a> <small>(${item.type})</small>`;
        resultsContainer.appendChild(div);
    });
}
"""
    with open(OUTPUT_DIR / "search.js", "w", encoding="utf-8") as f:
        f.write(js_content)


# ==============================================================================
# MAIN BUILD SYSTEM
# ==============================================================================
def build_website():
    print(f"Reading workbook: {INPUT_PATH}")
    xl = pd.ExcelFile(INPUT_PATH)

    sermons_df = xl.parse("Sermons") if "Sermons" in xl.sheet_names else pd.DataFrame()
    subjects_df = xl.parse("Subjects") if "Subjects" in xl.sheet_names else pd.DataFrame()
    suttas_df = xl.parse("Suttas") if "Suttas" in xl.sheet_names else pd.DataFrame()
    citations_df = xl.parse("Citations") if "Citations" in xl.sheet_names else pd.DataFrame()
    stats_df = xl.parse("Statistics") if "Statistics" in xl.sheet_names else pd.DataFrame()

    search_index = []

    sermon_count = len(sermons_df)
    subject_count = len(subjects_df)
    sutta_count = len(suttas_df)
    citation_count = len(citations_df)

    # 1. GENERATE HOMEPAGE
    home_content = f"""
    <div class="card">
        <h2>English Translations of Pahankanuwa Sermons</h2>
        <p><strong>Venerable Katukurunde Nanananda Thero</strong></p>
        <p>Welcome to the structured Dhamma Knowledge Base, cataloging and indexing the English translations of the Pahankanuwa Sermons delivered by Most Venerable Katukurunde Nanananda Thero. This portal provides direct cross-references between translated sermon passages, Dhamma subjects, and canonical Sutta citations.</p>
    </div>

    <div class="grid-4">
        <div class="stat-card">
            <div class="stat-number">{sermon_count}</div>
            <div class="stat-label">Translated Sermons</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{subject_count}</div>
            <div class="stat-label">Dhamma Subjects</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{citation_count}</div>
            <div class="stat-label">Citations Indexed</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{sutta_count}</div>
            <div class="stat-label">Sutta Sources</div>
        </div>
    </div>

    <div class="card">
        <h3>Knowledge Base Navigation</h3>
        <ul>
            <li><a href="sermons/index.html">Browse Translated Sermons</a> - Catalog of Pahankanuwa sermon translations and full DOCX transcripts</li>
            <li><a href="subjects/index.html">Browse Subjects</a> - Index of key Dhamma concepts and topics</li>
            <li><a href="suttas/index.html">Browse Sutta References</a> - Canonical Sutta index mapped to English translations</li>
            <li><a href="citations/index.html">Browse Citations Index</a> - Opening passages and textual citations</li>
            <li><a href="statistics.html">Workbook Statistics</a> - Detailed data and metrics breakdown</li>
        </ul>
    </div>
<div class="card">
    <h3>Knowledge Base Navigation</h3>
    <ul>
        <li><a href="sermons/index.html">Browse Sermons</a> - Catalog of Pahankanuwa sermon translations and full DOCX transcripts</li>
        <li>subjects/index.htmlBrowse Subjects</a> - Index of key Dhamma concepts and topics</li>
        <li>suttas/index.htmlBrowse Sutta References</a> - Canonical Sutta index mapped to English translations</li>
        <li>citations/index.htmlBrowse Citations Index</a> - Opening passages and textual citations</li>
        <li><a href="statistics.html">Workbook Statistics</a> - Detailed data and metrics breakdown</li>
    </ul>
</div>

<div class="card">
    <h3>Downloads</h3>

    <p>
        <a href="https://github.com/chamaniw/Pahankanuwa-english-translations/raw/refs/heads/main/Pahankanuwa_English_Translations.zip">
        📥 Download Complete Translation Collection (ZIP)
        </a>
    </p>

    <p>
        Download all available English translations of the Pahankanuwa sermon series in a single ZIP file.
    </p>
</div>
"""
    with open(OUTPUT_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Home", home_content))

    # 2. GENERATE SERMON PAGES
    sermon_index_rows = []
    for _, row in sermons_df.iterrows():
        sermon_id = str(row.get("Sermon ID", ""))
        opening_passage = str(row.get("Citation / Opening Passage", ""))
        canonical_source = str(row.get("Canonical Source (Likely Source)", ""))
        subjects = str(row.get("Dhamma Subject / Keywords", ""))
        notes = str(row.get("Notes", ""))

        sermon_name = format_sermon_docx_name(sermon_id)
        docx_url = f"{GITHUB_DOCX_BASE_URL}{sermon_name}.docx"

        filename = format_sermon_filename(sermon_id)
        page_url = f"sermons/{filename}"

        sermon_html = f"""
        <a href="index.html" class="back-link">&larr; Back to Translated Sermons List</a>
        <div class="card">
            <h2>Pahankanuwa Sermon: {sermon_id} (English Translation)</h2>
            <p><strong>Canonical Source:</strong> {canonical_source if canonical_source != 'nan' else 'N/A'}</p>
            <p><strong>Subjects / Keywords:</strong> {subjects if subjects != 'nan' else 'N/A'}</p>
        </div>

        <div class="card">
            <h3>Citation / Opening Passage</h3>
            <blockquote>{opening_passage if opening_passage != 'nan' else 'No opening citation listed.'}</blockquote>
        </div>

        <div class="card">
            <h3>Notes & Observations</h3>
            <p>{notes if notes != 'nan' else 'No additional notes.'}</p>
        </div>

        <div class="card">
            <h2>Resources</h2>
            <p><a href="{docx_url}" target="_blank" class="btn-docx">📄 Read Full Sermon Translation (DOCX)</a></p>
        </div>
        """
        with open(OUTPUT_DIR / "sermons" / filename, "w", encoding="utf-8") as f:
            f.write(get_base_html(f"Sermon {sermon_id}", sermon_html, rel_path_to_root="../"))

        sermon_index_rows.append(f"<tr><td><a href='{filename}'>{sermon_id}</a></td><td>{canonical_source}</td><td>{subjects}</td></tr>")
        search_index.append({"title": f"Sermon {sermon_id} (English)", "type": "Sermon Translation", "url": page_url, "text": f"{opening_passage} {subjects}"})

    # Sermon Index file
    sermons_list_html = f"""
    <h2>Pahankanuwa Sermons - English Translations Index</h2>
    <div class="card">
        <table>
            <thead><tr><th>Sermon ID</th><th>Canonical Source</th><th>Subjects</th></tr></thead>
            <tbody>{''.join(sermon_index_rows)}</tbody>
        </table>
    </div>
    """
    with open(OUTPUT_DIR / "sermons" / "index.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Translated Sermons Index", sermons_list_html, rel_path_to_root="../"))

    # 3. GENERATE SUBJECT PAGES
    subject_index_rows = []
    for _, row in subjects_df.iterrows():
        subj_name = str(row.get("Subject", "")).strip()
        sermons_linked = str(row.get("Sermons", ""))
        count = row.get("Count", "")

        if not subj_name or subj_name == 'nan':
            continue

        slug = slugify(subj_name)
        filename = f"{slug}.html"
        page_url = f"subjects/{filename}"

        linked_sermons_list = [s.strip() for s in sermons_linked.split(",") if s.strip()]
        sermon_links_html = "".join([f"<li><a href='../sermons/{format_sermon_filename(s)}'>Sermon {s} (English Translation)</a></li>" for s in linked_sermons_list])

        subj_html = f"""
        <a href="index.html" class="back-link">&larr; Back to Subjects List</a>
        <div class="card">
            <h2>Subject: {subj_name}</h2>
            <p><strong>Associated Sermon Count:</strong> {count}</p>
        </div>

        <div class="card">
            <h3>Related Sermon Translations</h3>
            <ul>{sermon_links_html if sermon_links_html else '<li>No sermons specifically listed.</li>'}</ul>
        </div>
        """
        with open(OUTPUT_DIR / "subjects" / filename, "w", encoding="utf-8") as f:
            f.write(get_base_html(f"Subject: {subj_name}", subj_html, rel_path_to_root="../"))

        subject_index_rows.append(f"<tr><td><a href='{filename}'>{subj_name}</a></td><td>{count}</td></tr>")
        search_index.append({"title": f"Subject: {subj_name}", "type": "Subject", "url": page_url, "text": subj_name})

    # Subjects Index File
    subjects_list_html = f"""
    <h2>Dhamma Subjects Index</h2>
    <div class="card">
        <table>
            <thead><tr><th>Subject</th><th>Translated Sermon Count</th></tr></thead>
            <tbody>{''.join(subject_index_rows)}</tbody>
        </table>
    </div>
    """
    with open(OUTPUT_DIR / "subjects" / "index.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Subjects Index", subjects_list_html, rel_path_to_root="../"))

    # 4. GENERATE SUTTA PAGES
    sutta_index_rows = []
    for _, row in suttas_df.iterrows():
        ref_name = str(row.get("Reference", "")).strip()
        sermons_linked = str(row.get("Sermons", ""))

        if not ref_name or ref_name == 'nan':
            continue

        slug = slugify(ref_name)
        filename = f"{slug}.html"
        page_url = f"suttas/{filename}"

        linked_sermons_list = [s.strip() for s in sermons_linked.split(",") if s.strip()]
        sermon_links_html = "".join([f"<li><a href='../sermons/{format_sermon_filename(s)}'>Sermon {s} (English Translation)</a></li>" for s in linked_sermons_list])

        sutta_html = f"""
        <a href="index.html" class="back-link">&larr; Back to Sutta Index</a>
        <div class="card">
            <h2>Sutta Reference: {ref_name}</h2>
        </div>

        <div class="card">
            <h3>Related Sermon Translations</h3>
            <ul>{sermon_links_html if sermon_links_html else '<li>No linked sermons listed.</li>'}</ul>
        </div>
        """
        with open(OUTPUT_DIR / "suttas" / filename, "w", encoding="utf-8") as f:
            f.write(get_base_html(f"Sutta: {ref_name}", sutta_html, rel_path_to_root="../"))

        sutta_index_rows.append(f"<tr><td><a href='{filename}'>{ref_name}</a></td></tr>")
        search_index.append({"title": f"Sutta: {ref_name}", "type": "Sutta", "url": page_url, "text": ref_name})

    # Sutta Index File
    suttas_list_html = f"""
    <h2>Suttas Index</h2>
    <div class="card">
        <table>
            <thead><tr><th>Sutta Reference</th></tr></thead>
            <tbody>{''.join(sutta_index_rows)}</tbody>
        </table>
    </div>
    """
    with open(OUTPUT_DIR / "suttas" / "index.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Suttas Index", suttas_list_html, rel_path_to_root="../"))

    # 5. GENERATE CITATION PAGES
    citation_index_rows = []
    for idx, row in citations_df.iterrows():
        citation_text = str(row.values[0]) if len(row) > 0 else f"Citation {idx+1}"
        sermons_linked = str(row.values[1]) if len(row) > 1 else ""

        slug = f"citation_{idx+1}"
        filename = f"{slug}.html"
        page_url = f"citations/{filename}"

        linked_sermons_list = [s.strip() for s in str(sermons_linked).split(",") if s.strip()]
        sermon_links_html = "".join([f"<li><a href='../sermons/{format_sermon_filename(s)}'>Sermon {s} (English Translation)</a></li>" for s in linked_sermons_list])

        citation_html = f"""
        <a href="index.html" class="back-link">&larr; Back to Citations Index</a>
        <div class="card">
            <h2>Citation #{idx+1}</h2>
            <p>{citation_text}</p>
        </div>

        <div class="card">
            <h3>Related Sermon Translations</h3>
            <ul>{sermon_links_html if sermon_links_html else '<li>No sermons linked directly.</li>'}</ul>
        </div>
        """
        with open(OUTPUT_DIR / "citations" / filename, "w", encoding="utf-8") as f:
            f.write(get_base_html(f"Citation {idx+1}", citation_html, rel_path_to_root="../"))

        citation_index_rows.append(f"<tr><td><a href='{filename}'>Citation #{idx+1}</a></td><td>{citation_text[:80]}...</td></tr>")
        search_index.append({"title": f"Citation #{idx+1}", "type": "Citation", "url": page_url, "text": citation_text})

    # Citations Index File
    citations_list_html = f"""
    <h2>Citations Index</h2>
    <div class="card">
        <table>
            <thead><tr><th>ID</th><th>Citation Excerpt</th></tr></thead>
            <tbody>{''.join(citation_index_rows)}</tbody>
        </table>
    </div>
    """
    with open(OUTPUT_DIR / "citations" / "index.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Citations Index", citations_list_html, rel_path_to_root="../"))

    # 6. GENERATE STATISTICS PAGE
    stats_table_html = ""
    if not stats_df.empty:
        stats_table_html = stats_df.to_html(classes="stats-table", index=False)
    else:
        stats_table_html = f"""
        <table>
            <tr><th>Metric</th><th>Count</th></tr>
            <tr><td>Total Translated Sermons</td><td>{sermon_count}</td></tr>
            <tr><td>Total Dhamma Subjects</td><td>{subject_count}</td></tr>
            <tr><td>Total Citations</td><td>{citation_count}</td></tr>
            <tr><td>Total Sutta Sources</td><td>{sutta_count}</td></tr>
        </table>
        """

    stats_page_html = f"""
    <h2>Knowledge Base Statistics</h2>
    <div class="card">
        {stats_table_html}
    </div>
    """
    with open(OUTPUT_DIR / "statistics.html", "w", encoding="utf-8") as f:
        f.write(get_base_html("Statistics", stats_page_html))

    # 7. WRITE SEARCH INDEX JSON
    with open(OUTPUT_DIR / "search_index.json", "w", encoding="utf-8") as f:
        json.dump(search_index, f)

    # 8. WRITE ASSETS
    write_assets()

    # PRINT SUMMARY OUTPUT
    print("\nWebsite generated successfully\n")
    print(f"Homepage:   {OUTPUT_DIR / 'index.html'}")
    print(f"Sermons:    {len(sermons_df)} pages generated in {OUTPUT_DIR / 'sermons'}")
    print(f"Subjects:   {len(subjects_df)} pages generated in {OUTPUT_DIR / 'subjects'}")
    print(f"Suttas:     {len(suttas_df)} pages generated in {OUTPUT_DIR / 'suttas'}")
    print(f"Citations:  {len(citations_df)} pages generated in {OUTPUT_DIR / 'citations'}")
    print(f"Statistics: {OUTPUT_DIR / 'statistics.html'}\n")

if __name__ == "__main__":
    build_website()