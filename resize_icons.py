from PIL import Image
import os

img = Image.open('extension/icon.png')
sizes = [16, 48, 128]

for size in sizes:
    resized = img.resize((size, size), Image.Resampling.LANCZOS)
    resized.save(f'extension/icons/icon{size}.png')

