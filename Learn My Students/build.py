#!/usr/bin/env python3
"""
Build script for Learn My Students project.
- Inlines all source files into learn-my-students.html
- Compiles README.md into a clean, beautifully formatted README.pdf (without browser headers/footers)
"""

import os
import json
import zipfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.join(SCRIPT_DIR, "src")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "learn-my-students.html")
SHARE_ZIP = os.path.join(SCRIPT_DIR, "Learn-My-Students.zip")

# Files that are safe to hand to colleagues. This list is deliberately explicit
# so no roster PDF or .deck.json (student data) can ever end up in the bundle.
SHARE_FILES = ["learn-my-students.html", "README.md", "README.pdf"]


def read_file(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def build_html():
    template = read_file(os.path.join(SRC_DIR, "index.template.html"))
    app_css = read_file(os.path.join(SRC_DIR, "app.css"))
    app_js = read_file(os.path.join(SCRIPT_DIR, "src", "app.js"))
    vendor_pdf = read_file(os.path.join(SRC_DIR, "vendor", "pdf.min.js"))
    vendor_worker = read_file(os.path.join(SRC_DIR, "vendor", "pdf.worker.min.js"))

    # Escape the worker source for embedding as a JS string literal.
    worker_escaped = json.dumps(vendor_worker)

    html = template
    html = html.replace("/* APP_CSS */", app_css)
    html = html.replace("/* VENDOR_PDF_JS */", vendor_pdf)
    html = html.replace("'/* VENDOR_PDF_WORKER_JS */'", worker_escaped)
    html = html.replace("/* APP_JS */", app_js)

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUTPUT_FILE) / 1024
    print(f"Built {OUTPUT_FILE} ({size_kb:.0f} KB)")


def build_pdf():
    try:
        from build_pdf import generate_pdf
        generate_pdf()
    except Exception as e:
        print(f"Warning: Could not build README.pdf ({e})")


def build_share_zip():
    """Bundle only the data-free, shareable files into a single zip."""
    included = []
    with zipfile.ZipFile(SHARE_ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in SHARE_FILES:
            path = os.path.join(SCRIPT_DIR, name)
            if os.path.exists(path):
                zf.write(path, arcname=name)
                included.append(name)
            else:
                print(f"Warning: {name} not found, skipping in share zip")
    size_kb = os.path.getsize(SHARE_ZIP) / 1024
    print(f"Packaged {os.path.basename(SHARE_ZIP)} ({size_kb:.0f} KB): {', '.join(included)}")


def build():
    build_html()
    build_pdf()
    build_share_zip()


if __name__ == "__main__":
    build()
