import os
import shutil
import re
from pathlib import Path
import json

try:
    import docx
except ImportError:
    docx = None

BASE_DIR = Path(r"C:\law-of-dependent-arising")
DOCS_DIR = BASE_DIR / "docs"
DOCS_AUDIO_DIR = DOCS_DIR / "audio"
DOCS_DOCX_DIR = DOCS_DIR / "docx"
DOCS_PDF_DIR = DOCS_DIR / "pdf"
DOCS_EPUB_DIR = DOCS_DIR / "epub"

DOCS_DIR.mkdir(parents=True, exist_ok=True)
DOCS_AUDIO_DIR.mkdir(parents=True, exist_ok=True)
DOCS_DOCX_DIR.mkdir(parents=True, exist_ok=True)
DOCS_PDF_DIR.mkdir(parents=True, exist_ok=True)
DOCS_EPUB_DIR.mkdir(parents=True, exist_ok=True)

SRC_DOCX_DIR = BASE_DIR / "docx"
SRC_AUDIO_DIR = BASE_DIR / "audio"
SRC_TXT_DIR = BASE_DIR / "txt"
SRC_PDF_DIR = BASE_DIR / "pdf"
SRC_EPUB_DIR = BASE_DIR / "epub"

def find_sermon_file(ch_num, format_ext, search_dir):
    if not search_dir.exists():
        return None
    seq_str = str(ch_num).zfill(3)
    for file in search_dir.iterdir():
        if file.is_file() and file.suffix.lower() == f".{format_ext}":
            if seq_str in file.name:
                return file.name
    return None

def get_audio_filename(ch_num):
    return f"Paticca_Samuppada_{str(ch_num).zfill(2)}.mp3"

def read_sermon_content_as_paragraphs(ch_num, s_num):
    paragraphs = []
    docx_filename = find_sermon_file(ch_num, "docx", SRC_DOCX_DIR)
    if docx_filename and (SRC_DOCX_DIR / docx_filename).exists():
        docx_path = SRC_DOCX_DIR / docx_filename
        if docx is not None:
            try:
                doc = docx.Document(docx_path)
                for p in doc.paragraphs:
                    text = p.text.strip()
                    if text:
                        paragraphs.append(text)
                if paragraphs:
                    return paragraphs
            except Exception as e:
                print("Error reading DOCX file:", e)

    txt_filename = find_sermon_file(ch_num, "txt", SRC_TXT_DIR)
    if txt_filename and (SRC_TXT_DIR / txt_filename).exists():
        try:
            with open(SRC_TXT_DIR / txt_filename, "r", encoding="utf-8") as f:
                content = f.read().strip()
                paras = [p.strip().replace("\n", " ") for p in content.split("\n\n") if p.strip()]
                if paras:
                    return paras
        except Exception as e:
            print("Error reading TXT file:", e)

    return [f"Text content for Sermon {s_num} is currently being prepared."]

def read_sermon_content_html(ch_num, s_num):
    paras = read_sermon_content_as_paragraphs(ch_num, s_num)
    return "".join(["<p>" + p.replace("\n", "<br>") + "</p>" for p in paras])

def normalize_text(text):
    import unicodedata
    if not text:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return text.lower()

SITE_CSS = r"""
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
    --reader-font-family: Georgia, "Times New Roman", serif;
    --reader-font-size: 1.15rem;
}

* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: Georgia, "Times New Roman", serif;
    background-color: var(--bg-color);
    color: var(--text-color);
    line-height: 1.8;
}

.site-header {
    background-color: var(--header-bg);
    color: var(--header-text);
    padding: 1.5rem 1rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    text-align: center;
}
.brand-section .logo-title {
    font-size: 1.8rem;
    font-weight: bold;
    color: #ffffff;
    margin-bottom: 0.2rem;
}
.brand-section .subtitle {
    font-size: 0.95rem;
    color: var(--header-subtitle);
    font-style: italic;
    padding: 0 0.5rem;
}

.top-nav {
    display: flex;
    justify-content: center;
    gap: 1rem;
    margin-top: 1.2rem;
    flex-wrap: wrap;
}
.top-nav a {
    color: #ffffff;
    text-decoration: none;
    padding: 0.5rem 0.9rem;
    border-radius: 4px;
    font-weight: 500;
    font-size: 0.95rem;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    transition: background-color 0.2s, color 0.2s;
}
.top-nav a:hover, .top-nav a.active {
    background-color: rgba(255, 255, 255, 0.25);
}

.main-content {
    max-width: 950px;
    margin: 2rem auto;
    padding: 0 1rem;
}

.hero {
    background: var(--card-bg);
    padding: 2.5rem 1.5rem;
    border-radius: 8px;
    border: 1px solid var(--border-color);
    margin-bottom: 2rem;
    box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    text-align: center;
}
.hero h1 { font-size: 2rem; color: #222; margin-bottom: 0.3rem; }
.hero h2 { font-size: 1.15rem; color: var(--primary-color); margin-bottom: 0.4rem; font-weight: normal; font-style: italic; }
.hero h3 { font-size: 1.2rem; color: #333; margin-bottom: 1.5rem; font-weight: normal; font-style: italic; }
.hero .author-title { font-size: 1.05rem; color: #555; margin-bottom: 2rem; font-style: italic; }

.pali-quote { 
    color: var(--header-bg); 
    text-align: center; 
    margin: 2.5rem 0; 
}

.button-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0, 1fr));
    gap: 0.8rem;
    margin: 2rem 0;
}
@media (max-width: 768px) {
    .button-grid {
        grid-template-columns: repeat(1, minmax(0, 1fr));
    }
    .main-content {
        margin: 1rem auto;
        padding: 0 0.5rem;
    }
    .hero {
        padding: 1.5rem 1rem;
    }
    .brand-section .logo-title {
        font-size: 1.5rem;
    }
}
.btn {
    display: inline-block;
    padding: 0.8rem 1rem;
    background-color: var(--header-bg);
    color: white !important;
    text-decoration: none;
    border-radius: 6px;
    font-weight: 600;
    font-size: 0.95rem;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    transition: background-color 0.2s;
    text-align: center;
}
.btn:hover { background-color: #5c0000; }
.btn-secondary {
    background-color: transparent;
    color: var(--header-bg) !important;
    border: 1px solid var(--header-bg);
}
.btn-secondary:hover { background-color: #f4eee6; }

.chapter-list {
    display: flex;
    flex-direction: column;
    gap: 1rem;
}
.chapter-card {
    background: var(--card-bg);
    border: 1px solid var(--border-color);
    border-radius: 6px;
    padding: 1.1rem 1.2rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
    transition: border-color 0.2s, transform 0.1s;
    text-decoration: none;
    color: inherit;
}
.chapter-card:hover {
    border-color: var(--header-bg);
    transform: translateY(-1px);
}
.chapter-title {
    font-size: 1.1rem;
    font-weight: bold;
    color: var(--header-bg);
}
.sermon-badge {
    font-size: 0.85rem;
    color: var(--text-muted);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

.search-input {
    width: 100%;
    padding: 0.9rem 1.2rem;
    font-size: 1.1rem;
    border: 1px solid var(--border-color);
    border-radius: 6px;
    background-color: #ffffff;
    box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    outline: none;
    font-family: Georgia, serif;
}
.search-input:focus {
    border-color: var(--header-bg);
    box-shadow: 0 0 0 3px rgba(128, 0, 0, 0.1);
}

.sermon-reader {
    background: #ffffff;
    border: 1px solid var(--border-color);
    border-radius: 8px;
    padding: 2.5rem 3rem;
    margin: 1.5rem 0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    font-family: var(--reader-font-family);
    font-size: var(--reader-font-size);
    line-height: 1.9;
}
@media (max-width: 768px) {
    .sermon-reader {
        padding: 1.2rem 1rem;
    }
}
.sermon-reader p {
    margin-bottom: 1.6rem;
    text-align: justify;
}
.sermon-nav-bar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin: 1.5rem 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
.sermon-nav-bar a {
    color: var(--header-bg);
    text-decoration: none;
    font-weight: 600;
}
.sermon-nav-bar a:hover { text-decoration: underline; }

.reader-controls {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background: #f4eee6;
    border: 1px solid var(--border-color);
    padding: 0.8rem 1.2rem;
    border-radius: 6px;
    margin-bottom: 1.5rem;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 0.9rem;
    flex-wrap: wrap;
    gap: 0.8rem;
}
.control-group {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    flex-wrap: wrap;
}
.control-group button, .download-btn {
    background: #fff;
    border: 1px solid var(--border-color);
    padding: 0.3rem 0.7rem;
    border-radius: 4px;
    cursor: pointer;
    font-weight: 600;
    color: var(--header-bg);
    text-decoration: none;
    font-size: 0.85rem;
    transition: background 0.2s, color 0.2s;
}
.control-group button:hover, .control-group button.active, .download-btn:hover {
    background: var(--header-bg);
    color: #fff;
    border-color: var(--header-bg);
}

mark {
    background-color: #ffeeba;
    color: inherit;
    padding: 0.1rem 0.2rem;
    border-radius: 2px;
    font-weight: bold;
}

footer {
    text-align: center;
    padding: 3rem 1rem;
    font-size: 0.9rem;
    color: var(--text-muted);
    border-top: 1px solid var(--border-color);
    margin-top: 4rem;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
"""

def build_header_html(active_tab="home"):
    tabs = [
        ("home", "Home", "index.html"),
        ("read", "Read", "read.html"),
        ("listen", "Listen", "listen.html"),
        ("search", "Search", "search.html"),
        ("about", "About", "about.html"),
    ]
    nav_links = []
    for tab_id, label, link in tabs:
        cls_attr = ' class="active"' if active_tab == tab_id else ''
        nav_links.append('<a href="' + link + '"' + cls_attr + '>' + label + '</a>')

    return (
        '<header class="site-header">'
        '<div class="brand-section">'
        '<div class="logo-title">The Law of Dependent Arising</div>'
        '<div class="subtitle">The Secret of Bondage and Release &mdash; Ven. Katukurunde Nanananda Thero</div>'
        '</div>'
        '<nav class="top-nav">' + "".join(nav_links) + '</nav>'
        '</header>'
    )

FOOTER_HTML = (
    '<footer>'
    '<p>This website is offered in the spirit of Dhamma Dana, freely available to all.</p>'
    '</footer>'
)

def generate_index_page():
    html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>The Law of Dependent Arising | Ven. Katukurunde Nanananda Thero</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('home') + r"""
<main class="main-content">
<section class="hero">
<h1>The Law of Dependent Arising</h1>
<h2>(Paṭicca Samuppāda)</h2>
<h3>The Secret of Bondage and Release</h3>
<div class="author-title">Ven. Katukurunde Ñāṇananda Thero</div>

<div class="pali-quote">
  <div style="font-style: italic; font-size: 1.15rem; line-height: 1.6; margin-bottom: 1.5rem;">
    “yo paṭicca samuppādaṁ passati<br>
    so Dhammaṁ passati<br><br>
    yo Dhammaṁ passati<br>
    so paṭicca samuppādaṁ passati”
  </div>
  <div style="font-style: italic; font-size: 1.05rem; font-weight: normal; color: #444; line-height: 1.6;">
    “He who sees Dependent Arising sees the Dhamma,<br>
    He who sees the Dhamma sees Dependent Arising.”
  </div>
</div>

<div class="button-grid">
<a href="read.html" class="btn">Read Online</a>
<a href="listen.html" class="btn">Listen</a>
<a href="search.html" class="btn">Search</a>
<a href="about.html" class="btn btn-secondary">About the Series</a>
</div>

<hr style="border: none; border-top: 1px solid var(--border-color); margin: 2.5rem 0;">

<div style="text-align: left;">
<h3 style="color: var(--header-bg); font-family: Georgia, serif; font-size: 1.3rem; margin-bottom: 1rem; font-style: normal;">Introduction</h3>
<p style="text-align: justify; line-height: 1.8;">
The Law of Dependent Arising: The Secret of Bondage and Release is the English translation of Ven. Katukurunde Nanananda Thero's twenty-sermon exposition of Paticca Samuppada (Dependent Arising), originally delivered as Sermons 183–202 of the Pahankanuwa sermon series. Inspired by a wish to clarify the profound teaching of Dependent Arising through the Buddha's own discourses, Ven. Nanananda devoted this series entirely to the themes of conditionality, bondage, and liberation.
</p>
</div>
</section>
</main>
""" + FOOTER_HTML + r"""
</body>
</html>"""
    with open(DOCS_DIR / "index.html", "w", encoding="utf-8") as f:
        f.write(html)

def generate_read_and_reader_pages():
    valid_sermons = [str(i).zfill(3) for i in range(183, 203)]

    chapter_cards = []
    for idx, s_num in enumerate(valid_sermons):
        ch_num = idx + 1
        card_html = (
            '<a href="chapter_' + str(ch_num) + '.html" class="chapter-card">'
            '<span class="chapter-title">Chapter ' + str(ch_num) + '</span>'
            '<span class="sermon-badge">Pahankanuwa Sermon ' + s_num + '</span>'
            '</a>'
        )
        chapter_cards.append(card_html)

    read_page_html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Read | The Law of Dependent Arising</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('read') + r"""
<main class="main-content">
<section class="hero" style="text-align: left; padding: 2rem 1.5rem;">
<h1 style="text-align: center; margin-bottom: 0.5rem; font-size: 1.8rem;">Book Chapters</h1>
<h2 style="text-align: center; margin-bottom: 2rem; color: var(--text-color); font-style: normal; font-size: 1.1rem;">The Law of Dependent Arising (Sermons 183–202)</h2>
<div class="chapter-list">
""" + "".join(chapter_cards) + r"""
</div>

<div style="margin-top: 2.5rem; background: #f4eee6; border: 1px solid var(--border-color); border-radius: 8px; padding: 1.5rem; text-align: center;">
<h3 style="color: var(--header-bg); font-family: Georgia, serif; font-size: 1.2rem; margin-bottom: 0.5rem;">Complete Book Edition (All 20 Sermons)</h3>
<p style="font-size: 0.95rem; color: var(--text-muted); margin-bottom: 1.2rem; font-family: -apple-system, sans-serif;">Download the complete unabridged collection in your preferred format:</p>
<div style="display: flex; justify-content: center; gap: 1rem; flex-wrap: wrap;">
<a href="Paticca_Samuppada_Sermons_Complete.docx" class="btn" style="padding: 0.6rem 1.2rem; font-size: 0.9rem;" download>Download Word (.docx)</a>
<a href="Paticca_Samuppada_Sermons_Complete.pdf" class="btn" style="padding: 0.6rem 1.2rem; font-size: 0.9rem;" download>Download PDF (.pdf)</a>
<a href="Paticca_Samuppada_Sermons_Complete.epub" class="btn" style="padding: 0.6rem 1.2rem; font-size: 0.9rem;" download>Download EPUB (.epub)</a>
</div>
</div>

</section>
</main>
""" + FOOTER_HTML + r"""
</body>
</html>"""
    with open(DOCS_DIR / "read.html", "w", encoding="utf-8") as f:
        f.write(read_page_html)

    for idx, s_num in enumerate(valid_sermons):
        ch_num = idx + 1
        prev_ch = "chapter_" + str(ch_num - 1) + ".html" if ch_num > 1 else None
        next_ch = "chapter_" + str(ch_num + 1) + ".html" if ch_num < len(valid_sermons) else None

        sermon_text = read_sermon_content_html(ch_num, s_num)
        
        docx_fname = find_sermon_file(ch_num, "docx", SRC_DOCX_DIR) or f"sermon_{s_num}_ENGLISH.docx"
        pdf_fname = find_sermon_file(ch_num, "pdf", SRC_PDF_DIR) or f"sermon_{s_num}_ENGLISH.pdf"
        epub_fname = find_sermon_file(ch_num, "epub", SRC_EPUB_DIR) or f"sermon_{s_num}_ENGLISH.epub"

        prev_link = '<a href="' + prev_ch + '">&larr; Previous Chapter</a>' if prev_ch else '<span></span>'
        next_link = '<a href="' + next_ch + '">Next Chapter &rarr;</a>' if next_ch else '<span></span>'

        reader_html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Chapter """ + str(ch_num) + r""" (Sermon """ + s_num + r""") | The Law of Dependent Arising</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('read') + r"""
<main class="main-content" style="max-width: 950px;">
<div class="sermon-nav-bar">
""" + prev_link + r"""
<a href="read.html">Table of Contents</a>
""" + next_link + r"""
</div>

<div class="reader-controls">
  <div class="control-group">
    <span>Font Size:</span>
    <button onclick="setFontSize('1rem')" id="size-sm">Small</button>
    <button onclick="setFontSize('1.15rem')" id="size-md" class="active">Medium</button>
    <button onclick="setFontSize('1.3rem')" id="size-lg">Large</button>
    <button onclick="setFontSize('1.5rem')" id="size-xl">X-Large</button>
  </div>
  <div class="control-group">
    <span>Font Type:</span>
    <button onclick="setFontFamily('Georgia, serif')" id="font-serif" class="active">Serif</button>
    <button onclick="setFontFamily('-apple-system, BlinkMacSystemFont, sans-serif')" id="font-sans">Sans</button>
  </div>
  <div class="control-group">
    <span>Download:</span>
    <a href="docx/""" + docx_fname + r"""" class="download-btn" download>DOCX</a>
    <a href="pdf/""" + pdf_fname + r"""" class="download-btn" download>PDF</a>
    <a href="epub/""" + epub_fname + r"""" class="download-btn" download>EPUB</a>
  </div>
</div>

<div style="text-align: center; margin-bottom: 2rem;">
<h1 style="font-size: 1.8rem; margin-bottom: 0.3rem;">Chapter """ + str(ch_num) + r"""</h1>
<p style="color: var(--text-muted); font-style: italic; font-family: -apple-system, sans-serif; font-size: 0.95rem;">(Pahankanuwa Sermon """ + s_num + r""")</p>
</div>

<article class="sermon-reader" id="readerArticle">
""" + sermon_text + r"""
</article>

<div class="sermon-nav-bar" style="margin-top: 2rem; border-top: 1px solid var(--border-color); padding-top: 1.5rem;">
""" + prev_link + r"""
<a href="read.html">Table of Contents</a>
""" + next_link + r"""
</div>
</main>
""" + FOOTER_HTML + r"""
<script>
function setFontSize(size) {
    const article = document.getElementById('readerArticle');
    article.style.fontSize = size;
    localStorage.setItem('preferred_font_size', size);
    document.querySelectorAll('.control-group button').forEach(b => {
        if(['Small','Medium','Large','X-Large'].some(t => b.innerText.includes(t))) {
            b.classList.remove('active');
        }
    });
    if(size === '1rem') document.getElementById('size-sm').classList.add('active');
    if(size === '1.15rem') document.getElementById('size-md').classList.add('active');
    if(size === '1.3rem') document.getElementById('size-lg').classList.add('active');
    if(size === '1.5rem') document.getElementById('size-xl').classList.add('active');
}

function setFontFamily(font) {
    const article = document.getElementById('readerArticle');
    article.style.fontFamily = font;
    localStorage.setItem('preferred_font_family', font);
    document.getElementById('font-serif').classList.toggle('active', font.includes('Georgia'));
    document.getElementById('font-sans').classList.toggle('active', !font.includes('Georgia'));
}

window.addEventListener('DOMContentLoaded', () => {
    const savedSize = localStorage.getItem('preferred_font_size');
    const savedFont = localStorage.getItem('preferred_font_family');
    if(savedSize) setFontSize(savedSize);
    if(savedFont) setFontFamily(savedFont);
});
</script>
</body>
</html>"""
        with open(DOCS_DIR / ("chapter_" + str(ch_num) + ".html"), "w", encoding="utf-8") as f:
            f.write(reader_html)

def generate_listen_page():
    audio_items = []
    for i in range(183, 203):
        s_num = str(i).zfill(3)
        ch_num = i - 182
        audio_name = get_audio_filename(ch_num)
        
        item_html = r"""
<div style="background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 1.5rem; margin-bottom: 1.5rem;">
<h3 style="color: var(--header-bg); margin-bottom: 0.5rem; font-family: Georgia, serif; font-size: 1.1rem;">Chapter """ + str(ch_num) + r""" (Sermon """ + s_num + r""")</h3>
<p style="font-size: 0.85rem; color: var(--text-muted); margin-bottom: 1rem; font-family: -apple-system, sans-serif;">English Audio Narration</p>
<audio controls style="width: 100%; margin-bottom: 0.8rem;">
<source src="audio/""" + audio_name + r"""" type="audio/mpeg">
Your browser does not support the audio element.
</audio>
<div>
<a href="audio/""" + audio_name + r"""" class="btn btn-secondary" style="font-size: 0.85rem; padding: 0.4rem 0.9rem;" download>Download MP3</a>
</div>
</div>"""
        audio_items.append(item_html)

    html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Listen | The Law of Dependent Arising</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('listen') + r"""
<main class="main-content">
<section class="hero" style="text-align: left; padding: 2rem 1.5rem;">
<h1 style="text-align: center; margin-bottom: 0.5rem; font-size: 1.8rem;">Audio Recordings</h1>
<h2 style="text-align: center; margin-bottom: 2rem; color: var(--text-color); font-style: normal; font-size: 1.1rem;">The Law of Dependent Arising Series</h2>
<p style="margin-bottom: 2rem; text-align: center; font-size: 0.95rem;">Listen online or download audio recordings for each sermon in the series.</p>
""" + "".join(audio_items) + r"""
</section>
</main>
""" + FOOTER_HTML + r"""
</body>
</html>"""
    with open(DOCS_DIR / "listen.html", "w", encoding="utf-8") as f:
        f.write(html)

def generate_search_page():
    search_list = []
    for i in range(183, 203):
        s_num = str(i).zfill(3)
        ch_num = i - 182
        title = f"Chapter {ch_num} (Sermon {s_num})"
        paragraphs = read_sermon_content_as_paragraphs(ch_num, s_num)
        
        search_item = {
            "chapter": ch_num,
            "sermon": s_num,
            "title": title,
            "paragraphs": paragraphs,
            "normalized_paragraphs": [normalize_text(p) for p in paragraphs]
        }
        search_list.append(search_item)
        
    search_json_str = json.dumps(search_list, ensure_ascii=False)

    html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Search | The Law of Dependent Arising</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('search') + r"""
<main class="main-content">
<section class="hero" style="padding: 2rem 1.5rem;">
<h1 style="margin-bottom: 0.5rem; font-size: 1.8rem;">Search the Book</h1>
<p style="text-align: center; margin-bottom: 2rem; font-size: 0.95rem;">Search all sermons, paragraphs, and Pali terms (diacritics are ignored automatically).</p>
<input type="text" id="searchBox" class="search-input" placeholder="Type to search (e.g., ignorance, paticca, feeling)...">
<div id="results" style="margin-top: 2rem; text-align: left;"></div>
</section>
</main>
""" + FOOTER_HTML + r"""
<script>
const searchData = """ + search_json_str + r""";

function normalize(str) {
    if (!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function performSearch() {
    const rawQuery = document.getElementById('searchBox').value.trim();
    const query = normalize(rawQuery);
    const results = document.getElementById('results');
    
    if (query.length < 2) {
        results.innerHTML = "";
        return;
    }
    
    let matchCount = 0;
    let resultsHtml = "";

    searchData.forEach(item => {
        item.normalized_paragraphs.forEach((normP, idx) => {
            if (normP.includes(query) && matchCount < 40) {
                matchCount++;
                
                const startIdx = Math.max(0, idx - 10);
                const endIdx = Math.min(item.paragraphs.length - 1, idx + 10);
                
                let passageHtml = "";
                for (let pIdx = startIdx; pIdx <= endIdx; pIdx++) {
                    let pText = item.paragraphs[pIdx];
                    let highlightedText = pText;
                    if (rawQuery.length > 0) {
                        try {
                            const termRegex = new RegExp('(' + escapeRegExp(rawQuery) + ')', 'gi');
                            highlightedText = pText.replace(termRegex, '<mark>$1</mark>');
                        } catch(e) {}
                    }

                    if (pIdx === idx) {
                        passageHtml += `<p style="margin-bottom: 0.8rem; background: #fdf6ec; padding: 0.4rem; border-left: 3px solid var(--header-bg);"><strong>[Paragraph ${pIdx + 1}]</strong> ${highlightedText}</p>`;
                    } else {
                        passageHtml += `<p style="margin-bottom: 0.8rem; color: #444;">${highlightedText}</p>`;
                    }
                }

                resultsHtml += `
                    <div style="background: var(--card-bg); border: 1px solid var(--border-color); border-radius: 8px; padding: 1.5rem; margin-bottom: 1.5rem; box-shadow: 0 2px 6px rgba(0,0,0,0.02);">
                      <h4 style="margin-bottom: 1rem; font-size: 1.15rem;">
                        <a href="chapter_${item.chapter}.html" style="color: var(--header-bg); text-decoration: none;">${item.title}</a>
                      </h4>
                      <div style="font-size: 0.95rem; line-height: 1.7; max-height: 400px; overflow-y: auto; padding-right: 0.5rem; border: 1px solid #eee; padding: 1rem; background: #faf8f5; border-radius: 6px;">
                        ${passageHtml}
                      </div>
                    </div>
                `;
            }
        });
    });

    if (matchCount === 0) {
        results.innerHTML = "<p style='color: var(--text-muted); text-align: center;'>No matching passages found.</p>";
    } else {
        results.innerHTML = `<p style="margin-bottom: 1rem; font-size: 0.9rem; color: var(--text-muted);">Found matching passages across the sermons:</p>` + resultsHtml;
    }
}

document.getElementById('searchBox').addEventListener('input', performSearch);
</script>
</body>
</html>"""
    with open(DOCS_DIR / "search.html", "w", encoding="utf-8") as f:
        f.write(html)

def generate_about_page():
    html = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>About | The Law of Dependent Arising</title>
<style>""" + SITE_CSS + r"""</style>
</head>
<body>
""" + build_header_html('about') + r"""
<main class="main-content">
<section class="hero" style="text-align: left; padding: 2.5rem 1.5rem;">
<h1 style="text-align: center; font-size: 1.8rem; margin-bottom: 0.5rem;">The Law of Dependent Arising</h1>
<h2 style="text-align: center; margin-bottom: 1.5rem; color: var(--text-color); font-style: normal; font-size: 1.1rem;">(Paticca Samuppada)<br>The Secret of Bondage and Release</h2>
<p style="text-align: center; font-style: italic; margin-bottom: 2rem; color: var(--primary-color); font-size: 0.95rem;">
Library Edition by Bhikkhu K. Nanananda<br>
Published by Katukurunde Nanananda Sadaham Senasun Bharaya, Sri Lanka
</p>
<h3 style="color: var(--header-bg); font-family: Georgia, serif; margin-top: 1.5rem; margin-bottom: 0.8rem; font-size: 1.2rem;">About the Book</h3>
<p style="margin-bottom: 1.2rem; line-height: 1.8;">
<em>The Law of Dependent Arising: The Secret of Bondage and Release</em> is the English translation of Ven. Katukurunde Nanananda Thero's twenty-sermon exposition of Paticca Samuppada (Dependent Arising), originally delivered as Sermons 183–202 of the Pahankanuwa sermon series.
</p>
<p style="margin-bottom: 1.2rem; line-height: 1.8;">
Inspired by a wish to clarify the profound teaching of Dependent Arising through the Buddha's own discourses, Ven. Nanananda devoted this series entirely to the themes of conditionality, bondage, and liberation. The work expands and develops many of the insights introduced in his renowned <em>Nibbana - The Mind Stilled</em> sermons and represents one of most sustained examinations of the Buddha's central teaching on dependent arising.
</p>
<h3 style="color: var(--header-bg); font-family: Georgia, serif; margin-top: 2rem; margin-bottom: 0.8rem; font-size: 1.2rem;">About Ven. Katukurunde Nanananda Thero</h3>
<p style="margin-bottom: 1.2rem; line-height: 1.8;">
Ven. Katukurunde Nanananda Mahathera (1940–2018) was a Sri Lankan Buddhist monk, Pali scholar, and meditator widely respected for his original and thought-provoking expositions of the Buddha's teachings.
</p>
<p style="margin-bottom: 1.2rem; line-height: 1.8;">
Among his best-known works are <em>Concept and Reality</em>, <em>Nibbana - The Mind Stilled</em>, and <i>The Law of Dependent Arising: The Secret of Bondage and Release</i>. His teachings place particular emphasis on direct insight into experience, dependent arising, emptiness, and the nature of liberation as presented in the early Buddhist texts.
</p>
<h3 style="color: var(--header-bg); font-family: Georgia, serif; margin-top: 2rem; margin-bottom: 0.8rem; font-size: 1.2rem;">Dhamma Dana</h3>
<blockquote style="border-left: 4px solid var(--header-bg); padding-left: 1rem; margin: 1.5rem 0; font-style: italic; color: #444; line-height: 1.7;">
"It is the express wish of Venerable Bhikkhu Katukurunde Nanananda that all his Dhamma Books and recorded sermons be offered as a pure gift of Dhamma free of charge to the Dhamma-thirsty world."
</blockquote>
<p style="margin-bottom: 1.2rem; line-height: 1.8;">
This website is offered in that same spirit of Dhamma Dana, freely available to all who wish to study, reflect upon, and practise the Buddha's teaching.
</p>
<h3 style="color: var(--header-bg); font-family: Georgia, serif; margin-top: 2rem; margin-bottom: 0.8rem; font-size: 1.2rem;">Source Material & Acknowledgements</h3>
<p style="margin-bottom: 0.8rem;">This website gratefully acknowledges:</p>
<ul style="margin-left: 1.5rem; margin-bottom: 1.5rem; line-height: 1.7;">
<li>The original Pahankanuwa sermon series, from which these sermons were delivered as Sermons 183–202.</li>
<li>The K.N.S.S.B. (Katukurunde Nanananda Sadaham Senasuna Buddhist Society) for publishing and preserving the English edition of <em>The Law of Dependent Arising: The Secret of Bondage and Release</em>.</li>
<li>Ven. Nanananda Thero and all those who contributed to the preservation and dissemination of these teachings.</li>
<li>The translators of the English edition. In the published volume, Ven. Nanananda records: <em>"I wish to thank Mrs. Lilian Ratnayake for translating on my behalf sermons No. 12, 13, 14, 16, 17, 18 and 19 in this volume. Due to failing health I could translate only sermons No. 1–11, 15 and 20."</em></li>
</ul>
</section>
</main>
""" + FOOTER_HTML + r"""
</body>
</html>"""
    with open(DOCS_DIR / "about.html", "w", encoding="utf-8") as f:
        f.write(html)

def generate_sitemap():
    base_url = "https://chamaniw.github.io/law-of-dependent-arising" # Replace with your GitHub Pages URL if needed
    urls = [
        ("/", "1.0", "daily"),
        ("/index.html", "1.0", "daily"),
        ("/read.html", "0.9", "weekly"),
        ("/listen.html", "0.8", "weekly"),
        ("/search.html", "0.8", "weekly"),
        ("/about.html", "0.5", "monthly"),
    ]
    
    for i in range(1, 21):
        urls.append((f"/chapter_{i}.html", "0.8", "monthly"))

    xml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'
    ]
    
    for path, priority, freq in urls:
        # If hosting locally or relative, you can also use relative or absolute URLs
        clean_path = path.lstrip('/')
        xml_lines.append('  <url>')
        xml_lines.append(f'    <loc>{base_url}/{clean_path}</loc>')
        xml_lines.append(f'    <changefreq>{freq}</changefreq>')
        xml_lines.append(f'    <priority>{priority}</priority>')
        xml_lines.append('  </url>')
        
    xml_lines.append('</urlset>')
    
    with open(DOCS_DIR / "sitemap.xml", "w", encoding="utf-8") as f:
        f.write("\n".join(xml_lines))

def copy_folder(src, dst):
    if src.exists() and src.is_dir():
        for f in src.iterdir():
            if f.is_file():
                shutil.copy2(f, dst / f.name)

def copy_complete_files():
    complete_files = [
        "Paticca_Samuppada_Sermons_Complete.docx",
        "Paticca_Samuppada_Sermons_Complete.epub",
        "Paticca_Samuppada_Sermons_Complete.pdf"
    ]
    for fname in complete_files:
        src_path = BASE_DIR / fname
        if src_path.exists():
            shutil.copy2(src_path, DOCS_DIR / fname)

def copy_assets():
    copy_folder(SRC_AUDIO_DIR, DOCS_AUDIO_DIR)
    copy_folder(SRC_DOCX_DIR, DOCS_DOCX_DIR)
    copy_folder(SRC_PDF_DIR, DOCS_PDF_DIR)
    copy_folder(SRC_EPUB_DIR, DOCS_EPUB_DIR)
    copy_complete_files()

if __name__ == "__main__":
    print("Building digital book website and XML sitemap...")
    generate_read_and_reader_pages()
    generate_listen_page()
    generate_search_page()
    generate_about_page()
    generate_index_page()
    generate_sitemap()
    copy_assets()
    print("Build complete successfully!")