import os
from win32com.client import Dispatch
from PIL import Image

def convert_pptx_to_images(pptx_file, output_folder):
    output_folder = os.path.abspath(output_folder)
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

    # Sort images based on the slide number extracted from the filename (e.g., slide_1.png, slide_2.png)
    images = sorted(images, key=lambda x: int(x.split('_')[-1].split('.')[0]))

    first_image = Image.open(images[0]).convert("RGB")
    image_list = [Image.open(img).convert("RGB") for img in images[1:]]
    first_image.save(output_pdf, save_all=True, append_images=image_list)

    # Cleanup: Remove all image files after PDF creation
    for img in images:
        try:
            os.remove(img)
        except OSError as e:
            print(f"Error deleting file {img}: {e}")

def convert_pptx_to_pdf(pptx_file):
    """
    Converts a PPTX file to a PDF, saving the PDF in the same folder with the same name.
    
    :param pptx_file: The path to the PPTX file.
    """
    # Ensure the input file is valid
    if not os.path.exists(pptx_file):
        raise FileNotFoundError(f"The file {pptx_file} does not exist.")
    
    # Get the folder and base filename of the input PPTX file
    folder = os.path.dirname(pptx_file)
    base_filename = os.path.splitext(os.path.basename(pptx_file))[0]
    
    # Define the output PDF path using the same folder and base filename
    output_pdf = os.path.join(folder, f"{base_filename}.pdf")
    
    # Step 1: Convert PPTX slides to images
    convert_pptx_to_images(pptx_file, folder)

    # Step 2: Convert images to a single PDF
    images_to_pdf(folder, output_pdf)

    print(f"PDF saved at {output_pdf}")
