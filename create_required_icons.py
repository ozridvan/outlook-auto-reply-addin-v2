from PIL import Image, ImageDraw
import os

def create_icon(size, filename):
    # Create image with Office blue background
    img = Image.new('RGB', (size, size), '#0078D4')
    draw = ImageDraw.Draw(img)
    
    # Add white rectangle for email icon
    margin = size // 6
    rect_width = size - 2 * margin
    rect_height = rect_width * 2 // 3
    
    x1 = margin
    y1 = (size - rect_height) // 2
    x2 = x1 + rect_width
    y2 = y1 + rect_height
    
    # Draw email envelope
    draw.rectangle([x1, y1, x2, y2], fill='white', outline='white')
    
    # Draw envelope flap
    center_x = size // 2
    flap_height = rect_height // 3
    draw.polygon([
        (x1, y1),
        (center_x, y1 + flap_height),
        (x2, y1)
    ], fill='#E0E0E0', outline='#E0E0E0')
    
    # Save the image
    img.save(filename, 'PNG')
    print(f"Created {filename}")

# Create the required icons for manifest validation
create_icon(64, 'assets/icon-64.png')   # Standard icon requirement
create_icon(128, 'assets/icon-128.png') # High resolution icon requirement

print("Required icons created successfully!")
