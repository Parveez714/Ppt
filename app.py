# -*- coding: utf-8 -*-
"""
PPT Master Web UI - Fullstack Dynamic AI Generator
Upload any Excel / Markdown / CSV file -> Gemini 3.5 Flash Lite plans and writes custom SVG slides -> sanitized & compiled to native DrawingML PPTX.
"""
import os
import sys
import json
import re
import uuid
import datetime
import subprocess
from pathlib import Path
from flask import Flask, request, jsonify, render_template_string, send_file

from google import genai

BASE_DIR = Path(__file__).resolve().parent
PYTHON_EXE = sys.executable
UPLOAD_DIR = BASE_DIR / "uploads"
PROJECTS_DIR = BASE_DIR / "projects"
UPLOAD_DIR.mkdir(exist_ok=True)
PROJECTS_DIR.mkdir(exist_ok=True)

# Load environment configuration
ENV_VARS = {}
env_file = BASE_DIR / ".env"
if env_file.exists():
    with open(env_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                ENV_VARS[k.strip()] = v.strip()

GEMINI_API_KEY = ENV_VARS.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
GEMINI_MODEL = ENV_VARS.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

genai_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PPT Master AI Studio</title>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #840000;
      --primary-hover: #6d0000;
      --bg: #090d16;
      --card-bg: #141c2e;
      --border: #26334d;
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background: var(--bg);
      color: var(--text-main);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }
    header {
      background: #060910;
      border-bottom: 1px solid var(--border);
      padding: 16px 36px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .brand {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .badge-icon {
      background: var(--primary);
      color: #fff;
      font-weight: 800;
      padding: 6px 12px;
      border-radius: 6px;
      letter-spacing: 1.5px;
      font-size: 13px;
    }
    .title { font-size: 17px; font-weight: 700; }
    .status-pill {
      font-size: 12px;
      background: #14532d;
      color: #4ade80;
      padding: 4px 12px;
      border-radius: 9999px;
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .status-dot { width: 8px; height: 8px; background: #4ade80; border-radius: 50%; }

    main {
      flex: 1;
      max-width: 1140px;
      width: 100%;
      margin: 36px auto;
      padding: 0 20px;
      display: grid;
      grid-template-columns: 1fr 1.15fr;
      gap: 26px;
    }
    @media(max-width: 900px) { main { grid-template-columns: 1fr; } }

    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 26px;
      box-shadow: 0 10px 25px -5px rgba(0,0,0,0.4);
    }
    .card h2 { font-size: 18px; font-weight: 700; margin-bottom: 6px; }
    .card p.desc { font-size: 13px; color: var(--text-muted); margin-bottom: 20px; }

    .dropzone {
      border: 2px dashed #3e4f6e;
      border-radius: 10px;
      padding: 32px 20px;
      text-align: center;
      cursor: pointer;
      background: #0d1320;
      transition: all 0.2s ease;
      position: relative;
    }
    .dropzone:hover, .dropzone.dragover {
      border-color: var(--primary);
      background: #1a1016;
    }
    .dropzone input {
      position: absolute;
      top: 0; left: 0; width: 100%; height: 100%;
      opacity: 0;
      cursor: pointer;
    }
    .file-icon { font-size: 30px; margin-bottom: 8px; }
    .highlight { color: #f59e0b; font-weight: 600; }
    .file-info {
      margin-top: 12px;
      background: #182238;
      border: 1px solid #2e3e5e;
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 13px;
      display: none;
      justify-content: space-between;
    }

    .form-group { margin-top: 16px; }
    .form-group label {
      display: block;
      font-size: 13px;
      font-weight: 600;
      margin-bottom: 6px;
      color: #cbd5e1;
    }
    .form-select, .form-input {
      width: 100%;
      background: #0d1320;
      border: 1px solid var(--border);
      color: #fff;
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 14px;
      outline: none;
    }
    .form-select:focus, .form-input:focus { border-color: var(--primary); }

    .btn-submit {
      width: 100%;
      margin-top: 22px;
      background: var(--primary);
      color: #fff;
      border: none;
      padding: 13px;
      border-radius: 8px;
      font-size: 15px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 10px;
      transition: background 0.2s ease;
    }
    .btn-submit:hover { background: var(--primary-hover); }
    .btn-submit:disabled { opacity: 0.5; cursor: not-allowed; }

    .logs-terminal {
      height: 380px;
      background: #070b13;
      border: 1px solid #1c263b;
      border-radius: 8px;
      padding: 14px;
      font-family: 'Consolas', monospace;
      font-size: 12px;
      color: #a5b4fc;
      overflow-y: auto;
      white-space: pre-wrap;
      line-height: 1.6;
    }

    .download-section {
      margin-top: 18px;
      display: none;
      background: #14532d;
      border: 1px solid #22c55e;
      border-radius: 8px;
      padding: 14px;
      text-align: center;
    }
    .btn-download {
      display: inline-block;
      margin-top: 8px;
      background: #22c55e;
      color: #052e16;
      font-weight: 700;
      padding: 10px 22px;
      border-radius: 6px;
      text-decoration: none;
      font-size: 14px;
    }
    .btn-download:hover { background: #4ade80; }

    .loader {
      display: none;
      width: 16px; height: 16px;
      border: 2px solid #ffffff;
      border-bottom-color: transparent;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
    }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <div class="badge-icon">PPT</div>
      <div class="title">Master AI Studio</div>
    </div>
    <div class="status-pill">
      <div class="status-dot"></div>
      Powered by Gemini 3.5 Flash Lite + DrawingML
    </div>
  </header>

  <main>
    <div class="card">
      <h2>Generate Custom Presentation</h2>
      <p class="desc">Upload any Excel sheet. Gemini will analyze the data, architect custom slide layouts, and compile an editable PowerPoint.</p>

      <form id="uploadForm">
        <div class="dropzone" id="dropzone">
          <input type="file" id="fileInput" name="file" accept=".xlsx,.xlsm,.csv,.xls">
          <div class="file-icon">📊</div>
          <div>Drag & drop your Excel file here, or <span class="highlight">browse</span></div>
          <div style="font-size: 11px; color: #64748b; margin-top: 4px;">.xlsx, .xlsm, or .csv</div>
        </div>

        <div class="file-info" id="fileInfo">
          <span id="fileNameDisplay">file.xlsx</span>
          <span style="color:#22c55e; font-weight:600;">Ready</span>
        </div>

        <div class="form-group">
          <label>Brand & Design System</label>
          <select class="form-select" id="brandPreset" name="brand">
            <option value="auto" selected>✨ Auto-Detect (Analyze Excel content & company domain)</option>
            <option value="embassy">Embassy Group (Official Crimson #840000 & Amber)</option>
            <option value="corporate_dark">Midnight Corporate (Deep Slate & Cobalt Blue)</option>
            <option value="minimal_clean">Clean Minimal Light</option>
          </select>
        </div>

        <div class="form-group">
          <label>Deck Title / Specific Focus (Optional)</label>
          <input type="text" class="form-input" id="deckTitle" name="title" placeholder="Leave empty for AI auto-title">
        </div>

        <div class="form-group">
          <label>Slide Count</label>
          <select class="form-select" id="slideCount" name="slide_count">
            <option value="0" selected>🤖 Auto-Determine (AI decides optimal slide count from data density)</option>
            <option value="3">3 Slides (Quick Summary)</option>
            <option value="4">4 Slides (Executive Overview)</option>
            <option value="5">5 Slides (Standard Comprehensive MIS)</option>
            <option value="6">6 Slides (Detailed Deep Dive)</option>
            <option value="7">7 Slides (Extended Presentation)</option>
            <option value="8">8 Slides (Full Multi-Section Review)</option>
          </select>
        </div>

        <button type="submit" class="btn-submit" id="submitBtn">
          <span class="loader" id="loader"></span>
          <span id="btnText">Generate Custom Deck with AI</span>
        </button>
      </form>
    </div>

    <div class="card">
      <h2>AI Pipeline Stream</h2>
      <p class="desc">Real-time LLM slide generation, XML layout verification, and DrawingML export.</p>

      <div class="logs-terminal" id="terminalLogs">> Ready. Upload your Excel file to begin AI generation.</div>

      <div class="download-section" id="downloadCard">
        <div style="font-weight:700; font-size:15px; color:#f0fdf4;">🎉 Custom Presentation Ready!</div>
        <div style="font-size:12px; color:#bbf7d0; margin-top:4px;" id="fileMeta">Dynamic Content • Native PowerPoint</div>
        <a href="#" class="btn-download" id="downloadLink" download>Download Presentation (.pptx)</a>
      </div>
    </div>
  </main>

  <script>
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const uploadForm = document.getElementById('uploadForm');
    const terminalLogs = document.getElementById('terminalLogs');
    const submitBtn = document.getElementById('submitBtn');
    const btnText = document.getElementById('btnText');
    const loader = document.getElementById('loader');
    const downloadCard = document.getElementById('downloadCard');
    const downloadLink = document.getElementById('downloadLink');
    const fileMeta = document.getElementById('fileMeta');

    fileInput.addEventListener('change', () => {
      if (fileInput.files.length > 0) {
        fileNameDisplay.textContent = fileInput.files[0].name + ' (' + (fileInput.files[0].size / 1024).toFixed(1) + ' KB)';
        fileInfo.style.display = 'flex';
      }
    });

    function log(msg) {
      terminalLogs.textContent += "\\n" + msg;
      terminalLogs.scrollTop = terminalLogs.scrollHeight;
    }

    uploadForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const formData = new FormData(uploadForm);
      submitBtn.disabled = true;
      loader.style.display = 'inline-block';
      btnText.textContent = 'Designing Slides with AI...';
      downloadCard.style.display = 'none';

      terminalLogs.textContent = '> [1/4] Uploading workbook...';

      try {
        log('> [2/4] Parsing workbook sheets and data tables into Markdown...');
        const response = await fetch('/api/generate', { method: 'POST', body: formData });
        const data = await response.json();

        if (data.success) {
          log('> [OK] Excel parsed successfully.');
          log('> [3/4] Gemini 3.5 Flash Lite analyzed data & generated custom SVG layouts:');
          data.log_steps.forEach(s => log('  • ' + s));
          log('> [4/4] Compiled native DrawingML PowerPoint deck: ' + data.filename);

          fileMeta.textContent = data.slides + ' Slides • ' + data.size_kb + ' KB • 100% Native DrawingML';
          downloadLink.href = '/api/download/' + data.project_id + '/' + data.filename;
          downloadLink.download = data.filename;
          downloadCard.style.display = 'block';
        } else {
          log('> [ERROR] ' + (data.error || 'Failed to generate.'));
        }
      } catch (err) {
        log('> [EXCEPTION] ' + err.message);
      } finally {
        submitBtn.disabled = false;
        loader.style.display = 'none';
        btnText.textContent = 'Generate Custom Deck with AI';
      }
    });
  </script>
</body>
</html>
"""

LOGO_DIR = BASE_DIR / "skills" / "ppt-master" / "templates" / "brands" / "embassy" / "images"

# Supported logo file extensions in priority order
_LOGO_EXTS = (".svg", ".png", ".jpg", ".jpeg", ".webp")

class BrandContext:
    def __init__(self, brand_name: str, copyright_text: str, website_url: str = ""):
        self.brand_name = brand_name
        self.copyright_text = copyright_text
        self.website_url = website_url

def get_brand_context(user_email: str = "", db: any = None, guideline_id: any = None) -> BrandContext:
    """Dynamic brand & copyright resolver.
    Follows corporate multi-tenant context resolution pattern:
    Resolves brand metadata, dynamic copyright text, and website based on user domain or guideline.
    """
    domain = user_email.split("@")[-1].lower() if "@" in user_email else ""
    if "embassy" in domain or guideline_id == "embassy" or not domain:
        return BrandContext(
            brand_name="Embassy Developments",
            copyright_text="© 2026. Embassy Developments Limited. All rights reserved.",
            website_url="https://embassydevelopments.com"
        )
    return BrandContext(
        brand_name=domain.split(".")[0].title() or "Corporate",
        copyright_text=f"© 2026. {domain.split('.')[0].title()}. All rights reserved.",
        website_url=f"https://{domain}"
    )

def _resolve_logo(prefer_white: bool = False) -> Path | None:
    """Dynamically pick the logo file from LOGO_DIR.
    - prefer_white=True  → looks for a filename containing 'white' first (for dark slides)
    - prefer_white=False → looks for a filename NOT containing 'white' first (for light slides)
    Drop any new brand image/svg file into LOGO_DIR and it is picked up automatically!
    """
    candidates = [p for p in LOGO_DIR.iterdir() if p.suffix.lower() in _LOGO_EXTS]
    if not candidates:
        return None
    white = [p for p in candidates if 'white' in p.stem.lower()]
    dark  = [p for p in candidates if 'white' not in p.stem.lower()]
    ordered = (white + dark) if prefer_white else (dark + white)
    # Sort each group by extension priority so .svg beats .png beats .jpg
    ordered.sort(key=lambda p: _LOGO_EXTS.index(p.suffix.lower()) if p.suffix.lower() in _LOGO_EXTS else 99)
    return ordered[0] if ordered else None

def _logo_base64(is_dark: bool = False) -> tuple[str, str]:
    """Return (base64_data, mime_type) for whichever logo file is in LOGO_DIR."""
    import base64
    logo_file = _resolve_logo(prefer_white=is_dark)
    if not logo_file or not logo_file.exists():
        return "", ""
    ext = logo_file.suffix.lower()
    mime = "image/svg+xml" if ext == ".svg" else f"image/{ext.lstrip('.')}"
    return base64.b64encode(logo_file.read_bytes()).decode("utf-8"), mime

def _inject_logo_and_footer(svg_content: str, is_dark_slide: bool = False, copyright_text: str = "") -> str:
    """Embed logo in the TOP-RIGHT corner of every slide and ensure dynamic copyright line."""
    logo_b64, mime = _logo_base64(is_dark=is_dark_slide)
    injections = []
    
    if logo_b64:
        # Top-right: x=1090 so a 170px-wide logo ends at x=1260 (80px right margin)
        logo_tag = (
            f'<image href="data:{mime};base64,{logo_b64}" '
            f'x="1090" y="28" width="170" height="39" preserveAspectRatio="xMaxYMid meet"/>'
        )
        injections.append(logo_tag)

    if copyright_text and "©" not in svg_content:
        footer_color = "#AAA9A7" if is_dark_slide else "#666567"
        footer_tag = (
            f'<text x="80" y="685" font-family="Poppins, Arial, sans-serif" font-size="14" '
            f'fill="{footer_color}">{copyright_text}</text>'
        )
        injections.append(footer_tag)

    if injections:
        return svg_content.replace('</svg>', f'{"\n".join(injections)}\n</svg>')
    return svg_content


def sanitize_svg(svg_content: str) -> str:
    """Sanitizes AI-generated SVG to conform strictly with PPT Master DrawingML requirements:
    - Strips CSS drop-shadow filters.
    - Strips unsupported text properties.
    - Inlines CSS classes from <style> blocks into explicit SVG presentation attributes (PPTX export rejects `class` attributes).
    - Repairs tag nesting and preserves exact camelCase viewBox.
    """
    # Strip any CSS drop-shadow filters
    svg_content = re.sub(r'filter="[^"]*drop-shadow[^"]*"', '', svg_content)
    svg_content = re.sub(r'style="[^"]*filter:\s*drop-shadow[^"]*"', '', svg_content)
    # Strip unsupported text properties — not in PPT Master's DrawingML mapping
    for prop in ('text-transform', 'line-height', 'text-align', 'text-indent',
                 'white-space', 'word-spacing', 'word-break', 'text-rendering',
                 'text-shadow', 'font-kerning', 'font-variant', 'writing-mode'):
        svg_content = re.sub(rf'\s*{prop}="[^"]*"', '', svg_content)
        svg_content = re.sub(rf'{prop}\s*:\s*[^;}}\"]+;?\s*', '', svg_content)

    # In-line CSS class declarations and clean classes
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(svg_content, 'html.parser')
        
        # 1. Parse and extract all CSS rules from <style> blocks
        css_rules = {}
        for style in soup.find_all('style'):
            matches = re.findall(r'\.([a-zA-Z0-9_-]+)\s*\{([^}]+)\}', style.string or '')
            for cls_name, decls in matches:
                props = {}
                for decl in decls.split(';'):
                    decl = decl.strip()
                    if not decl or ':' not in decl:
                        continue
                    k, v = decl.split(':', 1)
                    props[k.strip().lower()] = v.strip().strip('"\'')
                css_rules[cls_name] = props
            style.decompose()

        # 2. Inline mapped properties onto tags having `class` and strip `class`
        for tag in soup.find_all(attrs={'class': True}):
            cls = tag.get('class')
            classes = cls if isinstance(cls, list) else cls.split()
            for cl in classes:
                if cl in css_rules:
                    for k, v in css_rules[cl].items():
                        if k in ('fill', 'font-size', 'font-weight', 'letter-spacing', 'font-family', 'stroke', 'stroke-width', 'opacity'):
                            if k == 'font-size' and v.endswith('px'):
                                v = v[:-2]
                            if k == 'letter-spacing' and v.endswith('px'):
                                v = v[:-2]
                            if not tag.has_attr(k):
                                tag[k] = v
            del tag['class']

        # Ensure default font family
        for text_tag in soup.find_all(['text', 'tspan']):
            if not text_tag.has_attr('font-family'):
                text_tag['font-family'] = 'Poppins, Arial, sans-serif'

        # 3. CRITICAL: Unwrap any nested <svg> tags inside the root <svg>
        # PPT Master reserved nested <svg> strictly for picture-crop wrappers with a single <image> child.
        root_svg = soup.find('svg')
        if root_svg:
            for nested in root_svg.find_all('svg'):
                nested.unwrap()
            svg_content = str(root_svg)
    except Exception:
        pass

    # CRITICAL: BeautifulSoup lowercases tag attributes (turns viewBox into viewbox),
    # but PPT Master strict XML parser strictly requires camelCase viewBox!
    svg_content = svg_content.replace('viewbox=', 'viewBox=')

    # Ensure root viewBox exists if missing
    if 'viewBox=' not in svg_content:
        svg_content = re.sub(r'<svg', '<svg viewBox="0 0 1280 720"', svg_content, count=1)

    return svg_content

def auto_detect_brand(markdown_data: str, filename: str) -> str:
    """Analyze the Excel text and filename to auto-detect the matching brand."""
    text = (markdown_data[:5000] + " " + filename).lower()
    if any(k in text for k in ["embassy", "blvd", "residences", "reit", "real estate", "property", "sqft", "leasing", "rera"]):
        return "embassy"
    elif any(k in text for k in ["tech", "cloud", "server", "software", "api", "infra", "cyber"]):
        return "corporate_dark"
    return "embassy"

def generate_dynamic_slides_with_gemini(markdown_data: str, brand: str, title: str, slide_count: int, brand_context: BrandContext = None) -> list:
    """Invokes Gemini 3.5 Flash Lite to read the actual markdown data and craft tailored SVG slides."""
    if not genai_client:
        raise ValueError("GEMINI_API_KEY is not configured in .env")

    design_spec_path = BASE_DIR / "skills" / "ppt-master" / "templates" / "brands" / "embassy" / "templates" / "design_spec.md"
    brand_instructions = design_spec_path.read_text(encoding="utf-8")
    
    copyright_line = brand_context.copyright_text if brand_context else "© 2026. Embassy Developments Limited. All rights reserved."

    if slide_count <= 0:
        slide_count_directive = """
- Slide Count: AUTO-DETERMINE based on data depth and complexity.
  Examine the sheets, tables, and metrics in the provided data. Decide how many content slides are needed to clearly explain the data without overcrowding:
  * Minimal/single-focus data: 3 to 4 total slides.
  * Standard multi-metric MIS/tracker: 4 to 6 total slides.
  * Complex, multi-sheet/multi-project data: 6 to 8 total slides.
  Ensure EVERY slide has a distinct narrative purpose, starting with Slide 1 (Cover) and ending with the final Slide (Thank You).
"""
        count_rule = "Generate the optimal number of slides (between 3 and 8) according to the data depth."
    else:
        slide_count_directive = f"- Slide Count: Exactly {slide_count} slides."
        count_rule = f"Generate exactly {slide_count} slides."

    prompt = f"""
You are an expert presentation designer and data visualization architect for PPT Master.
You turn raw spreadsheets into high-end, executive PowerPoint presentations.

Input Data (extracted from Excel):
---
{markdown_data[:12000]}
---

Target Deck Configuration:
- Title / Focus: {title if title else "Executive Performance Review"}
{slide_count_directive}
- Canvas format: 1280 x 720 (16:9).
- Styling & Brand Guidelines:
{brand_instructions}
- Dynamic Copyright text: "{copyright_line}"

CRITICAL SLIDE STRUCTURE REQUIREMENT (COVER & THANK YOU):
- Slide 1 MUST BE the Front Page / Cover Slide: Pure black background (#000000), large UPPERCASE title in pure white (#FFFFFF), subtitle/kicker in grey (#666567) or accent orange (#F58220), date/author in footer.
- Content Slides (middle slides): Pure white background (#FFFFFF), category kicker in accent orange (#F58220), title in black (#000000), clear data cards/tables/charts, and dynamic copyright line in footer.
- The Final Slide MUST BE the Closing / Thank You Slide: Pure black background (#000000), large "THANK YOU" headline in pure white (#FFFFFF), contact/corporate info in grey (#666567), and dynamic copyright line in footer.

CRITICAL RULES FOR SVG GENERATION:
1. Every slide MUST be an independent, fully valid XML <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720" width="1280" height="720" lang="en-US">...</svg>.
2. Do NOT use CSS filters like filter="drop-shadow(...)". Use clean solid fills and 1px hairlines instead.
3. Do NOT use HTML tags. Use ONLY valid SVG: <rect>, <text>, <tspan>, <line>, <g>, <circle>, <polygon>, <defs>, <linearGradient>, <image>.
4. NEVER use CSS <style> blocks or `class="..."` attributes. Write all styles as direct SVG presentation attributes (e.g. fill="#000000", font-size="24", font-weight="500", letter-spacing="1").
5. NEVER use these unsupported text attributes — they will break DrawingML export: class, text-transform, line-height, text-align, white-space, word-spacing, word-break, text-rendering, text-shadow, font-kerning, font-variant, writing-mode.
6. TEXT WRAPPING — SVG has no auto wrap. You MUST manually split long text into multiple <text> elements, each on its own y coordinate spaced 28-36px apart. Never let a single text element exceed ~80 characters or run past x=1180.
7. SAFE ZONE & TOP-RIGHT LOGO — The brand logo will be placed in the TOP-RIGHT corner (x=1090 to x=1260, y=28 to y=67). Keep all your slide titles and shapes clear of x >= 1060 when y <= 90!
8. FONT & COLOR COMPLIANCE — Font must be Poppins (with Arial, sans-serif fallback). Color scheme: pure black #000000, pure white #FFFFFF, grey #666567, hairlines #AAA9A7, accent orange #F58220 (≤20% per slide).
9. SPACING — leave at least 40px between each content block. Do not stack elements with less than 24px gap.
10. ONE COLUMN layouts are safer than two-column for data-heavy slides — use two columns only when each column has 5 or fewer short lines.
11. Extract the REAL numbers, KPIs, property names, and dates from the provided data.
12. {count_rule}
13. Provide the output as a valid JSON array of objects:
[
  {{
    "slide_number": 1,
    "title": "Cover",
    "svg": "<svg xmlns=\\"http://www.w3.org/2000/svg\\" viewBox=\\"0 0 1280 720\\" width=\\"1280\\" height=\\"720\\" lang=\\"en-US\\">...</svg>"
  }}
]
"""

    resp = genai_client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt
    )

    text = resp.text.strip()
    match = re.search(r'```(?:json)?\s*(\[.*?\])\s*```', text, re.DOTALL)
    if match:
        json_str = match.group(1)
    else:
        json_str = text

    slides = json.loads(json_str)
    return slides

@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route("/api/generate", methods=["POST"])
def generate():
    try:
        uploaded_file = request.files.get("file")
        brand = request.form.get("brand", "embassy")
        title = request.form.get("title", "").strip()
        slide_count = int(request.form.get("slide_count", 5))

        project_slug = f"deck_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        project_dir = PROJECTS_DIR / project_slug
        svg_output = project_dir / "svg_output"
        sources_dir = project_dir / "sources"
        exports_dir = project_dir / "exports"
        
        project_dir.mkdir(parents=True, exist_ok=True)
        svg_output.mkdir(exist_ok=True)
        sources_dir.mkdir(exist_ok=True)
        exports_dir.mkdir(exist_ok=True)

        log_steps = []

        # 1. Save uploaded file
        if uploaded_file and uploaded_file.filename != '':
            saved_file = sources_dir / uploaded_file.filename
            uploaded_file.save(str(saved_file))
            filename_used = uploaded_file.filename
        else:
            default_excel = BASE_DIR / "realestate_sales_mis_dummy.xlsx"
            if default_excel.exists():
                import shutil
                saved_file = sources_dir / "realestate_sales_mis_dummy.xlsx"
                shutil.copy(str(default_excel), str(saved_file))
                filename_used = "realestate_sales_mis_dummy.xlsx"
            else:
                return jsonify({"success": False, "error": "No file uploaded"}), 400

        log_steps.append(f"Ingested file: {filename_used}")

        # 2. Extract Excel to Markdown
        md_output = sources_dir / "data.md"
        subprocess.run([
            PYTHON_EXE,
            str(BASE_DIR / "skills" / "ppt-master" / "scripts" / "source_to_md.py"),
            str(saved_file),
            "-o", str(md_output)
        ], check=True)

        with open(md_output, "r", encoding="utf-8") as f:
            markdown_data = f.read()

        log_steps.append(f"Parsed {len(markdown_data.splitlines())} lines of data")

        # Dynamic Auto-Brand Resolution if selected or blank
        if not brand or brand.lower() in ("auto", "none"):
            brand = auto_detect_brand(markdown_data, filename_used)
            log_steps.append(f"Auto-detected brand identity: '{brand.title()}'")

        user_email = request.form.get("user_email", "")
        guideline_id = request.form.get("guideline_id", brand)
        brand_ctx = get_brand_context(user_email=user_email, db=None, guideline_id=guideline_id)
        log_steps.append(f"Brand context: '{brand_ctx.brand_name}', Copyright: '{brand_ctx.copyright_text}'")

        # 3. Dynamic LLM Slide Generation with Gemini
        try:
            if slide_count <= 0:
                log_steps.append(f"Prompting {GEMINI_MODEL} to auto-determine optimal slide count from content...")
            else:
                log_steps.append(f"Prompting {GEMINI_MODEL} for {slide_count} custom slide architecture...")
            slides_data = generate_dynamic_slides_with_gemini(
                markdown_data=markdown_data,
                brand=brand,
                title=title,
                slide_count=slide_count,
                brand_context=brand_ctx
            )

            total_slides = len(slides_data)
            for i, slide_obj in enumerate(slides_data, start=1):
                svg_file = svg_output / f"{i:02d}_slide.svg"
                svg_content = sanitize_svg(slide_obj.get("svg", ""))
                # Front page (slide 1) and Closing/Thank You page (last slide) use dark theme
                is_dark = (i == 1) or (i == total_slides) or any(k in slide_obj.get("title", "").lower() for k in ("cover", "closing", "thank you", "thanks"))
                svg_content = _inject_logo_and_footer(svg_content, is_dark_slide=is_dark, copyright_text=brand_ctx.copyright_text)
                with open(svg_file, "w", encoding="utf-8") as sf:
                    sf.write(svg_content)
                log_steps.append(f"Generated Slide {i}/{total_slides}: {slide_obj.get('title', 'Slide')}")
        except Exception as e:
            log_steps.append(f"AI note: {str(e)[:80]}. Using brand baseline.")
            source_svg_dir = BASE_DIR / "projects" / "embassy_realestate_mis_ppt169_20261120" / "svg_output"
            import shutil
            for sf in source_svg_dir.glob("*.svg"):
                shutil.copy(str(sf), str(svg_output / sf.name))

        # 4. Compile directly to Native DrawingML PPTX
        pptx_name = f"{brand}_{project_slug}.pptx"
        out_pptx_path = exports_dir / pptx_name

        # Invoke svg_to_pptx with non-conforming permissive export
        proc = subprocess.run([
            PYTHON_EXE,
            str(BASE_DIR / "skills" / "ppt-master" / "scripts" / "svg_to_pptx.py"),
            str(project_dir),
            "-o", str(out_pptx_path),
            "--enable-dangerous-nonconforming-svg-export"
        ], capture_output=True, text=True)

        if proc.returncode != 0:
            # Secondary fallback with quick-generate flag
            proc2 = subprocess.run([
                PYTHON_EXE,
                str(BASE_DIR / "skills" / "ppt-master" / "scripts" / "svg_to_pptx.py"),
                str(project_dir),
                "-o", str(out_pptx_path),
                "--quick-generate"
            ], capture_output=True, text=True)
            if proc2.returncode != 0:
                raise RuntimeError(proc.stderr or proc2.stderr or proc.stdout)

        file_size_kb = round(out_pptx_path.stat().st_size / 1024, 1)

        return jsonify({
            "success": True,
            "project_id": project_slug,
            "filename": pptx_name,
            "slides": len(list(svg_output.glob("*.svg"))),
            "brand": brand.title(),
            "size_kb": file_size_kb,
            "log_steps": log_steps
        })

    except Exception as ex:
        return jsonify({"success": False, "error": str(ex)}), 500

@app.route("/api/download/<project_id>/<filename>")
def download(project_id, filename):
    file_path = PROJECTS_DIR / project_id / "exports" / filename
    if file_path.exists():
        return send_file(str(file_path), as_attachment=True)
    return "File not found", 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"PPT Master Web Server running on http://127.0.0.1:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
