from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Image as RLImage
from docx import Document
from docx.shared import Inches
import openpyxl
import io
import os

os.makedirs("outputs", exist_ok=True)

def image_to_pdf(image_bytes: bytes, filename: str) -> str:
    output_path = f"outputs/{filename}.pdf"
    image = Image.open(io.BytesIO(image_bytes))
    if image.mode != 'RGB':
        image = image.convert('RGB')
    temp_path = f"outputs/{filename}_temp.jpg"
    image.save(temp_path, "JPEG")

    page_width, page_height = letter
    margin = 50
    max_width = page_width - (margin * 2)
    max_height = page_height - (margin * 2)
    ratio = min(max_width / image.width, max_height / image.height)
    new_width = image.width * ratio
    new_height = image.height * ratio

    doc = SimpleDocTemplate(output_path, pagesize=letter,
        leftMargin=margin, rightMargin=margin,
        topMargin=margin, bottomMargin=margin)
    rl_image = RLImage(temp_path, width=new_width, height=new_height)
    doc.build([rl_image])
    return output_path

def image_to_word(image_bytes: bytes, filename: str) -> str:
    output_path = f"outputs/{filename}.docx"
    image = Image.open(io.BytesIO(image_bytes))
    if image.mode != 'RGB':
        image = image.convert('RGB')
    temp_path = f"outputs/{filename}_temp.jpg"
    image.save(temp_path, "JPEG")

    doc = Document()
    doc.add_heading('PHDCIScanner - Converted Image', 0)

    # Max width 6 inches, auto scale height
    max_width = 6.0
    aspect = image.height / image.width
    new_height = max_width * aspect
    # Max height 8 inches
    if new_height > 8.0:
        new_height = 8.0
        max_width = new_height / aspect

    doc.add_picture(temp_path, width=Inches(max_width))
    doc.save(output_path)
    return output_path

def image_to_excel(image_bytes: bytes, filename: str) -> str:
    output_path = f"outputs/{filename}.xlsx"
    image = Image.open(io.BytesIO(image_bytes))
    if image.mode != 'RGB':
        image = image.convert('RGB')

    # Resize image for Excel (max 800px wide)
    max_px = 800
    if image.width > max_px:
        ratio = max_px / image.width
        new_size = (int(image.width * ratio), int(image.height * ratio))
        image = image.resize(new_size, Image.LANCZOS)

    temp_path = f"outputs/{filename}_temp.jpg"
    image.save(temp_path, "JPEG")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Image"
    ws['A1'] = 'PHDCIScanner - Converted Image'
    img = openpyxl.drawing.image.Image(temp_path)
    img.anchor = 'A3'
    ws.add_image(img)
    wb.save(output_path)
    return output_path