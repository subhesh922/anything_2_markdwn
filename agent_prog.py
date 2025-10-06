import pymupdf4llm
import pathlib
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import fitz
from pdfminer.high_level import extract_text
from markitdown import MarkItDown
import os 
import subprocess
from docx import Document
import tempfile
from docx2pdf import convert
import win32com.client


def agent_file_processor(file_path):
    lower_path = file_path.lower()
    if lower_path.endswith('.pdf'):
        return ".pdf"
    elif lower_path.endswith('.txt'):
        return ".txt"
    elif lower_path.endswith('.docx'):
        return ".docx"
    elif lower_path.endswith('.csv'):
        return ".csv"
    elif lower_path.endswith('.xlsx'):
        return ".xlsx"
    elif lower_path.endswith('.pptx') or lower_path.endswith('.ppt'):
        return ".pptx"
    elif lower_path.endswith('.jpg') or lower_path.endswith('.jpeg') or lower_path.endswith('.png'):
        return ".png"
    else:
        return "Unsupported file type"    



def normal_pdf_processor(file_path):
    return pymupdf4llm.to_markdown(file_path)


def check_pdf_type(pdf_path):
    """
    Determine PDF type with more detailed classification.
   
    Returns:
        str: "text" (normal PDF), "scanned" (image-only),
             "hybrid" (mix of text and images), or "unknown"
    """
    try:
        import fitz
        from pdfminer.high_level import extract_text
       
        # First check with pdfminer
        text = extract_text(pdf_path)
        if len(text.strip()) > 100:  # Has significant text
            return "text"
           
        # Then check with PyMuPDF for images
        doc = fitz.open(pdf_path)
        has_images = any(page.get_image_info() for page in doc)
       
        if has_images and len(text.strip()) < 20:
            return "scanned"
        elif has_images and len(text.strip()) > 20:
            return "hybrid"
        else:
            return "unknown"
    except Exception as e:
        print(f"Analysis error: {e}")
        return "unknown"

def extract_text_to_markdown(
    pdf_path: str,
    lang: str = "eng",
    dpi: int = 300
) -> str:
    """
    Extracts text from a scanned PDF using PyMuPDF and pytesseract, then returns it as Markdown text using markitdown.

    Args:
        pdf_path (str): Path to the scanned PDF file.
        lang (str, optional): Language for OCR (e.g., "eng", "fra"). Default: "eng".
        dpi (int, optional): DPI for image conversion. Higher = better accuracy, slower. Default: 300.

    Returns:
        str: Markdown-formatted text extracted from the PDF.

    Raises:
        FileNotFoundError: If PDF file doesn't exist.
        Exception: If OCR or Markdown conversion fails.
    """
    try:
        fitz.TOOLS.mupdf_warnings(reset=False)  # <-- THIS IS THE KEY LINE
        # Step 1: Open the PDF with PyMuPDF
        doc = fitz.open(pdf_path)
        extracted_text = []

        # Step 2: Convert pages to images and perform OCR
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Render page as an image with specified DPI
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            # Convert pixmap to PIL Image for pytesseract
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            # Perform OCR
            text = pytesseract.image_to_string(img, lang=lang)
            extracted_text.append(text.strip())

        doc.close()

        # Step 3: Combine extracted text into a single string
        combined_text = "\n\n".join(extracted_text)

        # Step 4: Convert to Markdown using markitdown and capture output
        # Save to a temporary file since markitdown needs a file input
        temp_file = "temp_text.txt"
        with open(temp_file, "w", encoding="utf-8") as f:
            f.write(combined_text)

        # Run markitdown and capture stdout instead of saving to a file
        result = subprocess.run(
            ["markitdown", temp_file],
            capture_output=True,
            text=True,
            check=True
        )

        # Clean up temporary file
        os.remove(temp_file)

        # Return the Markdown text from markitdown's stdout
        return result.stdout

    except FileNotFoundError:
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    except Exception as e:
        raise Exception(f"Processing failed: {str(e)}")
    
def convert_docx_to_temp_pdf(docx_path: str) -> str:
    """
    Converts a .docx file to a temporary .pdf file.
    Returns the path to the temporary PDF (auto-deleted when closed).
    
    Args:
        docx_path (str): Path to the input .docx file.
    
    Returns:
        str: Path to the temporary PDF file.
    """
    # Create a temporary file (deleted when closed)
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        temp_pdf_path = temp_pdf.name
    
    # Convert DOCX to PDF (saved in temp file)
    convert(docx_path, temp_pdf_path)
    
    return temp_pdf_path

def ppt_to_pdf_win32com(ppt_path: str) -> str:
    """Convert PPT/PPTX to PDF using Microsoft PowerPoint (Windows only)"""
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    try:
        # Create temp PDF file
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
            pdf_path = temp_pdf.name
        
        # Convert
        deck = powerpoint.Presentations.Open(os.path.abspath(ppt_path))
        deck.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        deck.Close()
        return pdf_path
    finally:
        powerpoint.Quit()
def xlsx_to_mrkdwn(file_path):
    """
    Convert an Excel file to Markdown format.
    
    Args:
        file_path (str): Path to the Excel file.
        
    Returns:
        str: Markdown-formatted string.
    """
    try:
        md = MarkItDown()
        result = md.convert(file_path)
        return result.text_content
    except Exception as e:
        raise Exception(f"Error converting Excel to Markdown: {e}")
def csv_to_mrkdwn(file_path):
    """
    Convert an CSV file to Markdown format.
    
    Args:
        file_path (str): Path to the CSV file.
        
    Returns:
        str: Markdown-formatted string.
    """
    try:
        md = MarkItDown()
        result = md.convert(file_path)
        return result.text_content
    except Exception as e:
        raise Exception(f"Error converting CSV to Markdown: {e}")
def txt_to_mrkdwn(file_path):
    """
    Convert an TXT file to Markdown format.
    
    Args:
        file_path (str): Path to the TXT file.
        
    Returns:
        str: Markdown-formatted string.
    """
    try:
        md = MarkItDown()
        result = md.convert(file_path)
        return result.text_content
    except Exception as e:
        raise Exception(f"Error converting TXT to Markdown: {e}")
def extract_text_to_tempfile(image_path):
    """
    Extracts text from an image with preprocessing and saves to a temporary .txt file.
    
    Args:
        image_path: Path to the image file
        
    Returns:
        Path to the temporary text file containing extracted text,
        or None if extraction fails
    """
    try:
        # Open and preprocess the image
        with Image.open(image_path) as img:
            # Convert to grayscale
            img = img.convert('L')
            
            # Enhance contrast
            img = ImageEnhance.Contrast(img).enhance(2.0)
            
            # Binarize (thresholding)
            img = img.point(lambda x: 0 if x < 150 else 255)
            
            # Optional sharpening
            img = img.filter(ImageFilter.SHARPEN)
            
            # Perform OCR
            text = pytesseract.image_to_string(img).strip()
            
        # Create a temporary file
        with tempfile.NamedTemporaryFile(mode='w+', suffix='.txt', delete=False) as temp_file:
            temp_file.write(text)
            temp_path = temp_file.name
            
        return temp_path
    
    except Exception as e:
        print(f"OCR processing failed: {str(e)}")
        return None
