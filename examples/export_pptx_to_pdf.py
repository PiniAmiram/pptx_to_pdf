import os
from win32com.client import Dispatch
from PIL import Image

def convert_pptx_to_images(pptx_file, output_folder):
    # Convert output folder to an absolute path
    output_folder = os.path.abspath(output_folder)

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    powerpoint = Dispatch("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)

    for i, slide in enumerate(presentation.Slides, start=1):
        output_file = os.path.join(output_folder, f"slide_{i}.png")
        slide.Export(output_file, "PNG")
    
    presentation.Close()
    powerpoint.Quit()

def images_to_pdf(image_folder, output_pdf):
    images = sorted(
        [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(".png")]
    )
    if not images:
        raise FileNotFoundError("No PNG files found in the specified folder.")

    first_image = Image.open(images[0]).convert("RGB")
    image_list = [Image.open(img).convert("RGB") for img in images[1:]]
    first_image.save(output_pdf, save_all=True, append_images=image_list)

    # Cleanup: Remove all image files after PDF creation
    for img in images:
        try:
            os.remove(img)
        except OSError as e:
            print(f"Error deleting file {img}: {e}")

if __name__ == "__main__":
    import sys
    pptx_file = sys.argv[1]
    output_folder = sys.argv[2]
    pdf_file = sys.argv[3]  # Third argument for the output PDF file

    # Step 1: Convert PPTX slides to images
    convert_pptx_to_images(pptx_file, output_folder)

    # Step 2: Convert images to a single PDF
    images_to_pdf(output_folder, pdf_file)
