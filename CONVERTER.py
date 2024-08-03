import os
import tkinter as tk
from tkinter import filedialog
import comtypes.client
from pdf2image import convert_from_path
from PIL import Image
import sys

def docx_to_pdf_to_images_to_pdf(input_docx, output_dir ):
    # Convert Word to PDF
    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(input_docx))[0] + ".pdf")
    
    try:
        # Use absolute path
        abs_input_docx = os.path.abspath(input_docx)
        abs_pdf_path = os.path.abspath(pdf_path)
        
        print(f"Attempting to open: {abs_input_docx}")
        print(f"Attempting to save as: {abs_pdf_path}")
        
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(abs_input_docx)
        doc.SaveAs(abs_pdf_path, FileFormat=17)  # FileFormat=17 is for PDF
        doc.Close()
        word.Quit()
        print(f"Converted {abs_input_docx} to {abs_pdf_path}")
    except Exception as e:
        print(f"Error converting Word to PDF: {str(e)}")
        return None

    # Check if PDF was created
    if not os.path.exists(abs_pdf_path):
        print(f"PDF file was not created at {abs_pdf_path}")
        return None

    # Convert PDF to images
    try:
        print(f"Attempting to convert PDF to images using file: {abs_pdf_path}")
        print(f"PDF file size: {os.path.getsize(abs_pdf_path)} bytes")
        #print(f"Poppler path: {os.environ.get('PATH')}")
        
        images = convert_from_path(abs_pdf_path,500,poppler_path="./poppler-bin") # "C:/Program Files/poppler-24.02.0/Library/bin"
        image_paths = []
        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"page_{i+1}.jpg")
            image.save(image_path, 'JPEG')
            image_paths.append(image_path)
        print(f"Converted PDF to {len(image_paths)} images")
    except Exception as e:
        print(f"Error converting PDF to images: {str(e)}")
        #print(f"Python version: {sys.version}")
        #print(f"pdf2image version: {pdf2image.__version__}")
        return None

    # Convert images back to PDF
    try:
        final_pdf_path = os.path.join(output_dir, "final_output.pdf")
        images = [Image.open(image_path) for image_path in image_paths]
        images[0].save(final_pdf_path, save_all=True, append_images=images[1:])
        print(f"Converted images back to PDF: {final_pdf_path}")
        return final_pdf_path
    except Exception as e:
        print(f"Error converting images back to PDF: {str(e)}")
        return None

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    input_docx = filedialog.askopenfilename(
        title="Select Word document",
        filetypes=[("Word Document", "*.docx")]
    )
    
    if not input_docx:
        print("No file selected. Exiting.")
        return

    output_dir = filedialog.askdirectory(title="Select output directory")
    
    if not output_dir:
        print("No output directory selected. Exiting.")
        return

    if not os.path.exists(input_docx):
        print(f"Error: The file {input_docx} does not exist.")
        return

    if not os.path.isdir(output_dir):
        print(f"Error: The directory {output_dir} does not exist.")
        return

    try:
        final_pdf = docx_to_pdf_to_images_to_pdf(input_docx, output_dir)
        if final_pdf:
            print(f"Process completed successfully. Final PDF: {final_pdf}")
        else:
            print("Process failed to complete.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()