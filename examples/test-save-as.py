import os
from win32com.client import Dispatch
import comtypes.client

def save_pptx_as_pdf_comtypes(pptx_file):
        # Ensure the input file is valid
    if not os.path.exists(pptx_file):
        raise FileNotFoundError(f"The file {pptx_file} does not exist.")
    
    # Get the folder and base filename of the input PPTX file
    folder = os.path.dirname(pptx_file)
    base_filename = os.path.splitext(os.path.basename(pptx_file))[0]
    
    # Define the output PDF path using the same folder and base filename
    output_pdf = os.path.join(folder, f"{base_filename}.pdf")

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(pptx_file)
    presentation.SaveAs(output_pdf, 32)  # 32 is the format for PDF
    presentation.Close()
    powerpoint.Quit()

def convert_pptx_to_pdf_win32com(pptx_file):
    """
    Converts a PowerPoint presentation to a PDF using win32com.client.

    :param pptx_file: Path to the input PPTX file.
    """
    # Ensure the input file exists
    if not os.path.exists(pptx_file):
        print(f"Error: The input file '{pptx_file}' does not exist.")
        return
    
    # Get the folder and base filename of the input PPTX file
    folder = os.path.dirname(pptx_file)
    base_filename = os.path.splitext(os.path.basename(pptx_file))[0]
    
    # Define the output PDF path using the same folder and base filename
    output_pdf = os.path.join(folder, f"{base_filename}.pdf")

    # Create PowerPoint application object
    powerpoint = Dispatch("PowerPoint.Application")
    presentation = None

    try:
        # Open the PowerPoint presentation
        presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)
        
        # Save the presentation as a PDF
        presentation.SaveAs(output_pdf, 32)  # 32 is the format for PDF
        print(f"Converted {pptx_file} to {output_pdf}")
    except Exception as e:
        print(f"Error converting file: {e}")
    finally:
        # Close the presentation if it was opened successfully
        if presentation:
            presentation.Close()
        # Quit PowerPoint
        powerpoint.Quit()

if __name__ == "__main__":
  import sys
  if len(sys.argv) != 2:
      print("Usage: python script.py <pptx_file>")
      sys.exit(1)
  
  pptx_file = sys.argv[1]
  save_pptx_as_pdf_comtypes(pptx_file)
