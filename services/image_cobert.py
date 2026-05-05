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
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    page_width, page_height = letter
    ratio = min(page_width / image.width, page_height / image.height) * 0.9
    rl_image = RLImage(temp_path, width=image.width * ratio, height=image.height * ratio)
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
    doc.add_heading('PHDCIScanner - Image Document', 0)
    doc.add_picture(temp_path, width=Inches(6))
    doc.save(output_path)
    return output_path

def image_to_excel(image_bytes: bytes, filename: str) -> str:
    output_path = f"outputs/{filename}.xlsx"
    image = Image.open(io.BytesIO(image_bytes))
    if image.mode != 'RGB':
        image = image.convert('RGB')
    temp_path = f"outputs/{filename}_temp.jpg"
    image.save(temp_path, "JPEG")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Image"
    ws['A1'] = 'PHDCIScanner - Image'
    img = openpyxl.drawing.image.Image(temp_path)
    img.anchor = 'A3'
    ws.add_image(img)
    wb.save(output_path)
    return output_path