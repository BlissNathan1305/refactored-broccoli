# Install dependencies if needed:
# pip install pillow opencv-python noise numpy

from PIL import Image, ImageDraw, ImageFont
import numpy as np
import cv2
from noise import pnoise2
import random

# ---- Logo Settings ----
width, height = 800, 800  # Logo size
scale = 50.0
octaves = 5
persistence = 0.5
lacunarity = 2.0
seed = random.randint(0, 100)

# ---- Generate Perlin noise background ----
def generate_noise(width, height, scale, octaves, persistence, lacunarity, seed):
    noise_array = np.zeros((height, width))
    for y in range(height):
        for x in range(width):
            nx = x / scale
            ny = y / scale
            noise_val = pnoise2(nx, ny, octaves=octaves, persistence=persistence,
                                lacunarity=lacunarity, repeatx=1024, repeaty=1024, base=seed)
            noise_array[y][x] = noise_val
    # Normalize 0-255
    noise_array = ((noise_array - noise_array.min()) / (noise_array.max() - noise_array.min()) * 255).astype(np.uint8)
    return noise_array

# Generate noise
noise_img = generate_noise(width, height, scale, octaves, persistence, lacunarity, seed)

# Apply vibrant color gradient using OpenCV
color_maps = [cv2.COLORMAP_PLASMA, cv2.COLORMAP_JET, cv2.COLORMAP_MAGMA, cv2.COLORMAP_OCEAN]
colored_img = cv2.applyColorMap(noise_img, random.choice(color_maps))

# Optional: add Gaussian blur for smooth effect
blurred_img = cv2.GaussianBlur(colored_img, (7, 7), 0)

# Convert to Pillow Image
logo_img = Image.fromarray(cv2.cvtColor(blurred_img, cv2.COLOR_BGR2RGB))

# ---- Draw Logo Text or Symbol ----
draw = ImageDraw.Draw(logo_img)

# Add simple geometric symbol (circle in the center)
center_x, center_y = width // 2, height // 2
radius = width // 5
draw.ellipse((center_x-radius, center_y-radius, center_x+radius, center_y+radius),
             outline="white", width=8)

# Add text (logo initials)
try:
    font = ImageFont.truetype("arial.ttf", 80)  # You can use any font file
except:
    font = ImageFont.load_default()

draw.text((center_x - 60, center_y - 40), "AB", fill="white", font=font)

# ---- Save the logo ----
logo_img.save("generated_logo.png")
print("Logo saved as generated_logo.png")
