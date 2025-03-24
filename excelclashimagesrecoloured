import os
import zipfile
from PIL import Image

def extract_images_from_excel(input_excel_path, output_folder):
    """Extract all images from the Excel file and save them to a folder with their original names."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    with zipfile.ZipFile(input_excel_path, 'r') as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.startswith("xl/media/") and file_name.lower().endswith((".png", ".jpg", ".jpeg", ".bmp", ".gif")):
                normalized_name = os.path.basename(file_name).encode('utf-8', 'ignore').decode('utf-8')
                extracted_path = os.path.join(output_folder, normalized_name)
                if os.path.exists(extracted_path):
                    print(f"Skipping duplicate: {extracted_path}")
                    continue
                with open(extracted_path, "wb") as f:
                    f.write(zip_ref.read(file_name))
                print(f"Extracted: {extracted_path}")

def replace_colors(input_folder, output_folder):
    """Replace colors in images: green to blue and red to golden yellow."""
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
                    if g > r and g > b:
                        blue_intensity = int((g / 255) * 255)
                        pixels[i, j] = (0, 0, blue_intensity)

            # Second pass: Change red shades to shades of golden yellow
            for i in range(img.width):
                for j in range(img.height):
                    r, g, b = pixels[i, j]
                    if r > g and r > b:
                        yellow_intensity = int((r / 255) * 255)
                        pixels[i, j] = (yellow_intensity, int(yellow_intensity * 0.874), 0)

            output_path = os.path.join(output_folder, filename)
            img.save(output_path)
            print(f"Processed and saved: {output_path}")

# Set paths
input_excel_path = r"C:\Excel\TSA3.xlsx"
output_folder = r"C:\Excel\ExcelPics"
recolored_output_folder = r"C:\Excel\ExcelPicsRecolored"

# Extract images from the Excel file
extract_images_from_excel(input_excel_path, output_folder)

# Recolor the extracted images
replace_colors(output_folder, recolored_output_folder)
def replace_images_in_excel(input_excel_path, recolored_folder, output_excel_path):
    """Replace images in the Excel file with the recolored images and save as a new Excel file."""
    temp_folder = os.path.join(os.path.dirname(output_excel_path), "temp_excel")

    # Unzip the original Excel file to a temporary folder
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)
    with zipfile.ZipFile(input_excel_path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)

    # Replace images in the "xl/media" folder with recolored images
    media_folder = os.path.join(temp_folder, "xl", "media")
    if os.path.exists(media_folder):
        for filename in os.listdir(recolored_folder):
            recolored_image_path = os.path.join(recolored_folder, filename)
            target_image_path = os.path.join(media_folder, filename)
            if os.path.exists(target_image_path):
                os.remove(target_image_path)
            if os.path.exists(recolored_image_path):
                os.rename(recolored_image_path, target_image_path)
                print(f"Replaced: {target_image_path}")

    # Recompress the folder into a new Excel file
    with zipfile.ZipFile(output_excel_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for root, _, files in os.walk(temp_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, temp_folder)
                zip_ref.write(file_path, arcname)

    # Clean up the temporary folder
    for root, dirs, files in os.walk(temp_folder, topdown=False):
        for file in files:
            os.remove(os.path.join(root, file))
        for dir in dirs:
            os.rmdir(os.path.join(root, dir))
    os.rmdir(temp_folder)

    print(f"Recolored Excel file saved as: {output_excel_path}")

# Replace images in the Excel file
output_excel_path = r"C:\Excel\TSA2_recoloured.xlsx"
replace_images_in_excel(input_excel_path, recolored_output_folder, output_excel_path)
