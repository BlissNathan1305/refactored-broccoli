# Install dependencies if not installed:
# pip install pillow opencv-python noise numpy

from PIL import Image
import numpy as np
import cv2
from noise import pnoise2
import random

# ---- Wallpaper Settings ----
width, height = 2160, 3840  # 4K portrait (mobile-friendly)
num_wallpapers = 3           # Number of wallpapers to generate

# ---- Gradient Options ----
color_maps = [
    cv2.COLORMAP_PLASMA,
    cv2.COLORMAP_VIRIDIS,
    cv2.COLORMAP_JET,
    cv2.COLORMAP_OCEAN,
    cv2.COLORMAP_MAGMA,
    cv2.COLORMAP_TWILIGHT
]

# ---- Function to generate Perlin noise ----
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

# ---- Function to apply random gradient ----
def apply_random_gradient(noise_img):
    color_map = random.choice(color_maps)
    colored_img = cv2.applyColorMap(noise_img, color_map)
    return colored_img

# ---- Generate Wallpapers ----
for i in range(num_wallpapers):
    seed = random.randint(0, 100)
    scale = random.uniform(80, 150)  # Randomize noise scale
    octaves = random.randint(4, 8)
    persistence = random.uniform(0.4, 0.6)
    lacunarity = random.uniform(1.5, 2.5)

    # Generate noise
    noise_img = generate_noise(width, height, scale, octaves, persistence, lacunarity, seed)

    # Apply color gradient
    colored_img = apply_random_gradient(noise_img)

    # Random blur for smooth dreamy effect
    blur_value = random.choice([3,5,7,9])
    blurred_img = cv2.GaussianBlur(colored_img, (blur_value, blur_value), 0)

    # Convert to Pillow Image and save as 4K JPG
    final_img = Image.fromarray(cv2.cvtColor(blurred_img, cv2.COLOR_BGR2RGB))
    filename = f"mobile_wallpaper_4k_{i+1}.jpg"
    final_img.save(filename, "JPEG", quality=95)
    print(f"Wallpaper saved as {filename}")

print("All mobile 4K wallpapers generated successfully!")
