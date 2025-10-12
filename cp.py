from PIL import Image, ImageDraw, ImageFont
from docx.shared import Inches
import os

def create_cover_image():
    width, height = 2480, 3508  # A4 at 300 dpi
    bg_color_top = (25, 60, 120)
    bg_color_bottom = (60, 130, 200)

    # Create gradient background
    img = Image.new("RGB", (width, height), bg_color_top)
    draw = ImageDraw.Draw(img)
    for y in range(height):
        r = bg_color_top[0] + (bg_color_bottom[0] - bg_color_top[0]) * y // height
        g = bg_color_top[1] + (bg_color_bottom[1] - bg_color_top[1]) * y // height
        b = bg_color_top[2] + (bg_color_bottom[2] - bg_color_top[2]) * y // height
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # Title and author text
    title = "DATA ANALYSIS USING PYTHON AND R"
    subtitle = "A Practical Step-by-Step Guide"
    author = "Samuel Blessed Nathaniel"
    edition = "2025 Edition"

    try:
        # You can replace with any .ttf font available on your system
        title_font = ImageFont.truetype("DejaVuSans-Bold.ttf", 160)
        subtitle_font = ImageFont.truetype("DejaVuSans.ttf", 100)
        author_font = ImageFont.truetype("DejaVuSans-Italic.ttf", 80)
        edition_font = ImageFont.truetype("DejaVuSans.ttf", 70)
    except:
        title_font = ImageFont.load_default()
        subtitle_font = title_font
        author_font = title_font
        edition_font = title_font

    # Center text positions
    def center_text(text, y, font, fill=(255,255,255)):
        text_width, text_height = draw.textsize(text, font=font)
        x = (width - text_width) / 2
        draw.text((x, y), text, font=font, fill=fill)

    center_text(title, 1100, title_font)
    center_text(subtitle, 1400, subtitle_font)
    center_text(author, 2000, author_font)
    center_text(edition, 2200, edition_font)

    # Save cover
    cover_path = "cover_page.png"
    img.save(cover_path)
    return cover_path

def add_cover_to_doc(doc, cover_path):
    if os.path.exists(cover_path):
        doc.add_picture(cover_path, width=Inches(8.27), height=Inches(11.69))
        doc.add_page_break()
