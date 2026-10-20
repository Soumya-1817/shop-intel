"""
ShopIntel AI - Master Academic & Industrial Project Report Generator
Creates:
  1. ShopIntel_Project_Report.tex (Fully compilable LaTeX document)
  2. ShopIntel_Project_Report.docx (Formatted Microsoft Word document)

Adheres strictly to the requested Table of Contents:
  - Internship Certificate (i)
  - Declaration (ii)
  - Acknowledgement (iii)
  - Abstract (iv)
  - Contents (v)
  - List of Figures (vii)
  - Chapter 1: Introduction (1.1 - 1.5)
  - Chapter 2: Organization Profile & Internship Activities (2.1 - 2.4)
  - Chapter 3: Technology Stack & System Architecture (3.1 - 3.4)
  - Chapter 4: Key Modules & Features (4.1 - 4.5)
  - Chapter 5: User Interface Walkthrough & Implementation (5.1 - 5.2)
  - Chapter 6: Testing, Results & Observations (6.1 - 6.4)
  - Chapter 7: Future Scope (7.1)
  - Chapter 8: Conclusion (8.1 - 8.2)
  - References
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

print("Loading report generation pipeline...")
