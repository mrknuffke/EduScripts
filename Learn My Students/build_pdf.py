#!/usr/bin/env python3
"""
Build script to compile README.md into a beautifully formatted README.pdf.
Eliminates default browser header/footer and ensures optimal page breaks & visual hierarchy.
"""

import os
import subprocess
import re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
README_MD = os.path.join(SCRIPT_DIR, "README.md")
README_PDF = os.path.join(SCRIPT_DIR, "README.pdf")
TEMP_HTML = os.path.join(SCRIPT_DIR, "readme_build_temp.html")

CSS_STYLING = """<style>
@page {
  size: letter;
  margin: 14mm 15mm 14mm 15mm;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
  color: #1e293b;
  line-height: 1.45;
  font-size: 9.5pt;
  margin: 0;
  padding: 0;
}

h1 {
  font-size: 18.5pt;
  color: #0f172a;
  margin-top: 0;
  margin-bottom: 6px;
  font-weight: 700;
  border-bottom: 2px solid #e2e8f0;
  padding-bottom: 6px;
}

/* Subtitle blockquote */
h1 + blockquote {
  background: #f8fafc;
  border-left: 4px solid #2563eb;
  margin: 0 0 12px 0;
  padding: 8px 12px;
  border-radius: 0 6px 6px 0;
  color: #334155;
  font-size: 10pt;
}

h1 + blockquote p {
  margin: 0;
}

h2 {
  font-size: 12pt;
  color: #0f172a;
  margin-top: 14px;
  margin-bottom: 6px;
  font-weight: 600;
  page-break-after: avoid;
  break-after: avoid;
}

/* Force Flashcards section to start cleanly at top of page 2 if needed */
#how-the-flashcards-work-leitner-5-box-system {
  page-break-before: always;
  break-before: page;
}

h3 {
  font-size: 10pt;
  color: #1e293b;
  margin-top: 10px;
  margin-bottom: 4px;
  font-weight: 600;
  page-break-after: avoid;
  break-after: avoid;
}

hr {
  border: 0;
  height: 1px;
  background: #e2e8f0;
  margin: 12px 0;
  page-break-after: avoid;
  break-after: avoid;
}

p {
  margin-top: 0;
  margin-bottom: 6px;
  color: #334155;
}

ul, ol {
  padding-left: 20px;
  margin-top: 4px;
  margin-bottom: 6px;
  page-break-inside: avoid;
  break-inside: avoid;
}

li {
  margin-bottom: 3px;
  color: #334155;
}

code {
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
  font-size: 8.8pt;
  background-color: #f1f5f9;
  color: #0f172a;
  padding: 1px 4px;
  border-radius: 4px;
  border: 1px solid #e2e8f0;
}

/* Warning blockquote */
blockquote {
  background-color: #fffbe6;
  border-left: 4px solid #d97706;
  margin: 8px 0;
  padding: 8px 12px;
  border-radius: 0 6px 6px 0;
  page-break-inside: avoid;
  break-inside: avoid;
}

blockquote p {
  margin: 0;
  color: #92400e;
}

/* FAQ Details boxes */
details {
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  border-radius: 6px;
  padding: 8px 12px;
  margin-bottom: 6px;
  page-break-inside: avoid;
  break-inside: avoid;
}

summary {
  font-weight: 600;
  color: #0f172a;
  margin-bottom: 2px;
}

details br {
  display: none;
}

/* Table styling */
table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 8px;
  margin-bottom: 10px;
  font-size: 8.8pt;
  page-break-inside: avoid;
  break-inside: avoid;
}

th, td {
  border: 1px solid #cbd5e1;
  padding: 6px 10px;
}

th {
  background-color: #f1f5f9;
  color: #0f172a;
  font-weight: 600;
}

tr:nth-child(even) {
  background-color: #f8fafc;
}
</style>
"""

def generate_pdf():
    res = subprocess.run(
        ["/opt/homebrew/bin/pandoc", README_MD, "-f", "markdown", "-t", "html"],
        capture_output=True,
        text=True,
        check=True
    )
    body_html = res.stdout

    # Add 'open' attribute to <details> tags so FAQ items are expanded in PDF
    body_html = re.sub(r'<details>', r'<details open>', body_html)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Student Name Learner - README</title>
{CSS_STYLING}
</head>
<body>
{body_html}
</body>
</html>
"""

    with open(TEMP_HTML, "w", encoding="utf-8") as f:
        f.write(full_html)

    chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    cmd = [
        chrome_path,
        "--headless=new",
        "--no-pdf-header-footer",
        f"--print-to-pdf={README_PDF}",
        TEMP_HTML
    ]

    subprocess.run(cmd, check=True)

    if os.path.exists(TEMP_HTML):
        os.remove(TEMP_HTML)

    size_kb = os.path.getsize(README_PDF) / 1024
    print(f"Successfully generated {README_PDF} ({size_kb:.1f} KB)")

if __name__ == "__main__":
    generate_pdf()
