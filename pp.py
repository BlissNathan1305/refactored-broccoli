def create_graphic_cover(path="cover_page.png"):
    """
    Generates a modern, visually appealing cover page for the data analysis book.
    Works with Pillow >= 10.0 (uses textbbox instead of deprecated textsize).
    """
    from PIL import Image, ImageDraw, ImageFont
    import random

    # A4 300DPI
    width, height = 2480, 3508  

    # Gradient background (blue → teal)
    top_color = (25, 60, 120)
    bottom_color = (40, 160, 180)

    img = Image.new("RGB", (width, height), top_color)
    draw = ImageDraw.Draw(img)

    # Vertical gradient
    for y in range(height):
        r = top_color[0] + (bottom_color[0] - top_color[0]) * y // height
        g = top_color[1] + (bottom_color[1] - top_color[1]) * y // height
        b = top_color[2] + (bottom_color[2] - top_color[2]) * y // height
        draw.line([(0, y), (width, y)], fill=(r, g, b))

    # Decorative geometric pattern (light overlay lines)
    for x in range(0, width, 250):
        draw.line([(x, 0), (x - 500, height)], fill=(255, 255, 255, 20), width=3)

    # Add faint data nodes (circles)
    for _ in range(100):
        cx = random.randint(0, width)
        cy = random.randint(0, height)
        radius = random.randint(5, 15)
        draw.ellipse((cx - radius, cy - radius, cx + radius, cy + radius),
                     fill=(255, 255, 255, random.randint(20, 60)))

    # Text details
    title = "DATA ANALYSIS USING PYTHON AND R"
    subtitle = "A Comprehensive Practical Guide"
    author = "Samuel Blessed Nathaniel"
    edition = "2025 Edition"

    # Load fonts (fallback-safe)
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

    # Helper function for centering text
    def center_text(line, y, font, fill=(255, 255, 255)):
        bbox = draw.textbbox((0, 0), line, font=font)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (width - w) / 2
        draw.text((x, y), line, font=font, fill=fill)

    # Positioning the texts beautifully
    center_text(title, 1100, title_font)
    center_text(subtitle, 1400, subtitle_font, fill=(230, 255, 255))
    center_text(author, 2000, author_font, fill=(255, 255, 240))
    center_text(edition, 2200, edition_font, fill=(220, 240, 255))

    img.save(path)
    return path
