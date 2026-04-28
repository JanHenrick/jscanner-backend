from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os

def convert_text_to_pdf(text: str, filename: str) -> str:
    output_path = f"outputs/{filename}.pdf"
    c = canvas.Canvas(output_path, pagesize=letter)
    
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 750, "PHDCIScanner - Converted Document")
    
    c.setFont("Helvetica", 11)
    y = 710
    for line in text.split('\n'):
        if line.strip():
            c.drawString(50, y, line[:90])
            y -= 20
            if y < 50:
                c.showPage()
                c.setFont("Helvetica", 11)
                y = 750
    
    c.save()
    return output_path