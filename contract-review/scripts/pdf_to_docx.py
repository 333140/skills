#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to DOCX converter for contract review workflow.

Converts PDF contracts to DOCX format so they can be processed by the
contract-review workflow (which requires .docx for comment annotation).

Usage:
    python pdf_to_docx.py input.pdf output.docx

Dependencies:
    - pdfplumber (text extraction)
    - python-docx (DOCX generation)

Exit codes:
    0 - Success
    1 - File not found
    2 - Invalid input
    3 - Processing error
"""

import sys
import argparse
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber not installed. Run: pip install pdfplumber", file=sys.stderr)
    sys.exit(3)

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx", file=sys.stderr)
    sys.exit(3)


def pdf_to_docx(pdf_path: str, docx_path: str) -> bool:
    """Convert PDF to DOCX preserving paragraph structure."""
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        print(f"Error: PDF not found: {pdf_path}", file=sys.stderr)
        return False

    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = '仿宋'
    font.size = Pt(12)

    with pdfplumber.open(str(pdf_file)) as pdf:
        print(f"Processing {len(pdf.pages)} pages...")

        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue

            # Split into paragraphs and add to document
            paragraphs = text.split('\n')
            for para_text in paragraphs:
                stripped = para_text.strip()
                if not stripped:
                    continue

                # Try to detect headings
                if stripped.startswith('第') and ('章' in stripped or '条' in stripped):
                    p = doc.add_heading(stripped, level=2)
                elif len(stripped) < 30 and ('：' in stripped or ':' in stripped):
                    # Likely a label/title line
                    p = doc.add_paragraph(stripped)
                    p.runs[0].bold = True
                else:
                    p = doc.add_paragraph(stripped)

            # Add page break between pages (except last)
            if page_num < len(pdf.pages):
                doc.add_page_break()

    doc.save(docx_path)
    print(f"Converted: {pdf_path} -> {docx_path}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF contract to DOCX for review",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s contract.pdf contract.docx
  %(prog)s contract.pdf contract.docx
        """
    )

    parser.add_argument("input", help="Input PDF file")
    parser.add_argument("output", help="Output DOCX file")

    args = parser.parse_args()

    success = pdf_to_docx(args.input, args.output)
    sys.exit(0 if success else 3)


if __name__ == "__main__":
    main()
