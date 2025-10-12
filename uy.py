"""
Data Analysis Using Python and R - Auto Book Generator
Author: Samuel Blessed Nathaniel
Version: 2025 Edition
Generates a professionally formatted .docx book with a designed cover page.
"""

from docx import Document
from docx.shared import Inches
from PIL import Image, ImageDraw, ImageFont
import os, random, textwrap


# =========================================================
# 1. CREATE BEAUTIFUL GRAPHIC COVER
# =========================================================
def create_graphic_cover(path="cover_page.png"):
    """
    Generates a modern, visually appealing cover page.
    Compatible with Pillow >= 10.0 (uses textbbox).
    """
    width, height = 2480, 3508  # A4 at 300 DPI
    top_color = (25, 60, 120)
    bottom_color = (40, 160, 180)

    img = Image.new("RGB", (width, height), top_color)
    draw = ImageDraw.Draw(img)

    # Gradient background
    for y in range(height):
        r = top_color[0] + (bottom_color[0] - top_color[0]) * y // height
        g = top_color[1] + (bottom_color[1] - top_color[1]) * y // height
        b = top_color[2] + (bottom_color[2] - top_color[2]) * y // height
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # Decorative diagonal lines
    for x in range(0, width, 250):
        draw.line([(x, 0), (x - 500, height)], fill=(255, 255, 255, 30), width=3)

    # Data node circles
    for _ in range(120):
        cx = random.randint(0, width)
        cy = random.randint(0, height)
        r = random.randint(6, 15)
        draw.ellipse(
            (cx - r, cy - r, cx + r, cy + r),
            fill=(255, 255, 255, random.randint(25, 70)),
        )

    # Text
    title = "DATA ANALYSIS USING PYTHON AND R"
    subtitle = "A Comprehensive Practical Guide"
    author = "Samuel Blessed Nathaniel"
    edition = "2025 Edition"

    # Fonts (use defaults if custom fonts unavailable)
    try:
        title_font = ImageFont.truetype("DejaVuSans-Bold.ttf", 160)
        subtitle_font = ImageFont.truetype("DejaVuSans.ttf", 100)
        author_font = ImageFont.truetype("DejaVuSans-Italic.ttf", 80)
        edition_font = ImageFont.truetype("DejaVuSans.ttf", 70)
    except:
        title_font = ImageFont.load_default()
        subtitle_font = title_font
        author_font = title_font
        edition_font = title_font

    # Helper for centering text
    def center_text(line, y, font, fill=(255, 255, 255)):
        bbox = draw.textbbox((0, 0), line, font=font)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (width - w) / 2
        draw.text((x, y), line, font=font, fill=fill)

    # Write the texts
    center_text(title, 1100, title_font)
    center_text(subtitle, 1400, subtitle_font, fill=(230, 255, 255))
    center_text(author, 2000, author_font, fill=(255, 255, 240))
    center_text(edition, 2200, edition_font, fill=(220, 240, 255))

    img.save(path)
    return path


# =========================================================
# 2. GENERATE BOOK CONTENT
# =========================================================
def generate_book_content():
    """Returns structured chapters for the book"""
    chapters = [
        ("Introduction to Data Analysis", """
Data analysis is the process of inspecting, cleaning, transforming, and modeling data 
to discover useful information, draw conclusions, and support decision-making.

Python and R are the two most popular languages for data analysis due to their powerful libraries, 
ease of use, and large communities. Python is excellent for automation and machine learning, 
while R excels in statistical modeling and visualization.
        """),

        ("Data Collection and Importing", """
Data can come from multiple sources: spreadsheets, databases, APIs, or text files.
In Python, pandas makes data import easy using `pd.read_csv()` or `pd.read_excel()`.
In R, you can use `read.csv()` or the `readxl` package for Excel files.
        """),

        ("Data Cleaning and Preparation", """
Data cleaning ensures accuracy and consistency. Typical tasks include:
- Handling missing values (`dropna`, `fillna` in pandas or `na.omit` in R)
- Removing duplicates
- Correcting data types
- Standardizing formats

Python Example:
