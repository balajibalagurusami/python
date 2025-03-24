import os
from PIL import Image

def replace_colors(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            img_path = os.path.join(input_folder, filename)
            img = Image.open(img_path).convert("RGB")
            pixels = img.load()

            # First pass: Change green shades to shades of blue
            for i in range(img.width):
                for j in range(img.height):
                    r, g, b = pixels[i, j]

                    # Check if green is dominant
                    if g > r and g > b:
                        # Scale blue intensity based on green intensity
                        blue_intensity = int((g / 255) * 255)
                        pixels[i, j] = (0, 0, blue_intensity)

            # Second pass: Change red shades to shades of golden yellow
            for i in range(img.width):
                for j in range(img.height):
                    r, g, b = pixels[i, j]

                    # Check if red is dominant
                    if r > g and r > b:
                        # Scale yellow intensity based on red intensity
                        yellow_intensity = int((r / 255) * 255)
                        pixels[i, j] = (yellow_intensity, int(yellow_intensity * 0.874), 0)  # (R, G, B) for golden yellow

            output_path = os.path.join(output_folder, filename)
            img.save(output_path)
            print(f"Processed and saved: {output_path}")

# Example usage
input_folder = r"C:\Clash\APP\TSA3"  # Replace with your input folder path
output_folder = r"C:\Clash\APP\TSA3\recoloured"  # Replace with your output folder path
replace_colors(input_folder, output_folder)
