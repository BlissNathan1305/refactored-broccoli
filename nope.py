# Install dependencies if not already installed:
# pip install pillow opencv-python noise numpy

from PIL import Image
import numpy as np
import cv2
from noise import pnoise2
import random

# Wallpaper settings
width, height = 1920, 1080  # Full HD
scale = 100.0               # Noise scale
octaves = 6
persistence = 0.5
lacunarity = 2.0
seed = random.randint(0, 100)

# Generate Perlin noise array
def generate_noise(width, height, scale, octaves, persistence, lacunarity, seed):
    noise_array = np.zeros((height, width))
    for y in range(height):
        for x in range(width):
            nx = x / scale
            ny = y / scale
            noise_val = pnoise2(nx, ny, octaves=octaves, persistence=persistence, 
                                lacunarity=lacunarity, repeatx=1024, repeaty=1024, base=seed)
            noise_array[y][x] = noise_val
    # Normalize to 0-255
    noise_array = ((noise_array - noise_array.min()) / (noise_array.max() - noise_array.min()) * 255).astype(np.uint8)
    return noise_array

# Generate noise
noise_img = generate_noise(width, height, scale, octaves, persistence, lacunarity, seed)

# Apply color gradient
def apply_gradient(noise_img):
    # Convert grayscale to color
    gradient_img = cv2.applyColorMap(noise_img, cv2.COLORMAP_PLASMA)
    return gradient_img

colored_img = apply_gradient(noise_img)

# Add Gaussian blur for smooth effect
blurred_img = cv2.GaussianBlur(colored_img, (9, 9), 0)

# Convert to Pillow Image and save as JPG
final_img = Image.fromarray(cv2.cvtColor(blurred_img, cv2.COLOR_BGR2RGB))
final_img.save("stunning_wallpaper.jpg", "JPEG", quality=95)

print("Wallpaper generated and saved as stunning_wallpaper.jpg")
