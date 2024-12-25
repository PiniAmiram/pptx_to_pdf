import os
from win32com.client import Dispatch
from PIL import Image

def convert_pptx_to_images(pptx_file, output_folder):
    # Convert output folder to an absolute path
    output_folder = os.path.abspath(output_folder)

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Initialize PowerPoint application
    powerpoint = Dispatch("PowerPoint.Application")

    # Open the PowerPoint presentation
    presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)

    # Store the original slide dimensions
    original_width = presentation.PageSetup.SlideWidth
    original_height = presentation.PageSetup.SlideHeight

    # Increase slide dimensions for higher resolution export
    presentation.PageSetup.SlideWidth = original_width * 2  # Double the width
    presentation.PageSetup.SlideHeight = original_height * 2  # Double the height

    # Loop through each slide, scale content, and export as PNG
    for i, slide in enumerate(presentation.Slides, start=1):
        # Scale content to fit the new slide size
        scale_slide_content(slide, original_width * 2, original_height * 2)

        # Export the slide as a PNG file
        output_file = os.path.join(output_folder, f"slide_{i}.png")
        slide.Export(output_file, "PNG")

    # Reset the slide dimensions to the original size
    presentation.PageSetup.SlideWidth = original_width
    presentation.PageSetup.SlideHeight = original_height

    # Close the presentation
    presentation.Close()
    powerpoint.Quit()

def scale_slide_content(slide, new_width, new_height):
    # Loop through all shapes on the slide and scale them proportionally
    for shape in slide.Shapes:
        if shape.HasTextFrame:
            # Scale text box size proportionally
            shape.Width = shape.Width * (new_width / slide.Parent.PageSetup.SlideWidth)
            shape.Height = shape.Height * (new_height / slide.Parent.PageSetup.SlideHeight)
            shape.Left = shape.Left * (new_width / slide.Parent.PageSetup.SlideWidth)
            shape.Top = shape.Top * (new_height / slide.Parent.PageSetup.SlideHeight)
        elif shape.Type == 4:  # Shape type 4 is an image
            # Scale images proportionally
            shape.LockAspectRatio = True
            shape.Width = shape.Width * (new_width / slide.Parent.PageSetup.SlideWidth)
            shape.Height = shape.Height * (new_height / slide.Parent.PageSetup.SlideHeight)
            shape.Left = shape.Left * (new_width / slide.Parent.PageSetup.SlideWidth)
            shape.Top = shape.Top * (new_height / slide.Parent.PageSetup.SlideHeight)

def images_to_pdf(image_folder, output_pdf):
    # Get all PNG files in the folder and sort them by slide index
    images = sorted(
        [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(".png")],
        key=lambda x: int(x.split('_')[-1].split('.')[0])  # Sort based on slide number
    )
    
    if not images:
        raise FileNotFoundError("No PNG files found in the specified folder.")

    # Open the first image and convert it to RGB
    first_image = Image.open(images[0]).convert("RGB")
    
    # Open the rest of the images and convert them to RGB
    image_list = [Image.open(img).convert("RGB") for img in images[1:]]
    
    # Save the images as a single PDF
    first_image.save(output_pdf, save_all=True, append_images=image_list)

    # Cleanup: Remove all image files after PDF creation
    for img in images:
        try:
            os.remove(img)
        except OSError as e:
            print(f"Error deleting file {img}: {e}")

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Usage: python script.py <pptx_file> <output_folder> <output_pdf>")
        sys.exit(1)
    
    pptx_file = sys.argv[1]
    output_folder = sys.argv[2]
    pdf_file = sys.argv[3]  # Third argument for the output PDF file

    # Step 1: Convert PPTX slides to images with increased resolution
    convert_pptx_to_images(pptx_file, output_folder)

    # Step 2: Convert images to a single PDF
    images_to_pdf(output_folder, pdf_file)
