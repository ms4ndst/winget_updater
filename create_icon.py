from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    # Create images of different sizes for the ico file
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []
    
    for size in sizes:
        # Create a new image with a transparent background
        image = Image.new('RGBA', size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        
        # Calculate dimensions
        width, height = size
        padding = width // 8
        
        # Create background circle
        circle_bbox = [padding, padding, width - padding, height - padding]
        draw.ellipse(circle_bbox, fill=(0, 120, 212))  # Windows blue color
        
        # Add text
        font_size = int(width / 2)
        try:
            font = ImageFont.truetype("arial.ttf", font_size)
        except:
            font = ImageFont.load_default()
            
        # Draw the "W"
        text = "W"
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        
        x = (width - text_width) // 2
        y = (height - text_height) // 2 - height // 10  # Slight upward adjustment
        
        draw.text((x, y), text, fill="white", font=font)
        images.append(image)
    
    # Save as ICO file
    icon_path = "app.ico"
    images[0].save(icon_path, format='ICO', sizes=[(size[0], size[1]) for size in sizes])
    return icon_path

if __name__ == "__main__":
    icon_path = create_icon()
    print(f"Icon created successfully at: {icon_path}")

