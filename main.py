from fastapi import FastAPI, UploadFile, File 
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import os

app = FastAPI(
    title="PHDCIScanner - SnapConvert AI",
    description="Scan images and convert to Word, PDF, Excel",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create outputs folder if not exists
os.makedirs("outputs", exist_ok=True)

@app.get("/")
def root():
    return {"message": "PHDCIScanner API is running!"}

@app.get("/health")
def health():
    return {"status": "OK", "version": "1.0.0"}

@app.post("/convert/text")
async def convert_to_text(file: UploadFile = File(...)):
    contents = await file.read()
    from services.ocr_service import extract_text_from_image
    text = extract_text_from_image(contents)
    return {"filename": file.filename, "text": text}

@app.post("/convert/word")
async def convert_to_word(file: UploadFile = File(...)):
    contents = await file.read()
    from services.ocr_service import extract_text_from_image
    from services.word_export import convert_text_to_word
    
    text = extract_text_from_image(contents)
    output_path = convert_text_to_word(text, file.filename)
    
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="converted.docx"
    )

@app.post("/convert/pdf")
async def convert_to_pdf(file: UploadFile = File(...)):
    contents = await file.read()
    from services.ocr_service import extract_text_from_image
    from services.pdf_export import convert_text_to_pdf
    
    text = extract_text_from_image(contents)
    output_path = convert_text_to_pdf(text, file.filename)
    
    return FileResponse(
        output_path,
        media_type="application/pdf",
        filename="converted.pdf"
    )
    
@app.post("/convert/excel")
async def convert_to_excel(file: UploadFile = File(...)):
    contents = await file.read()
    from services.ocr_service import extract_text_from_image
    from services.excel_export import convert_text_to_excel
    
    text = extract_text_from_image(contents)
    output_path = convert_text_to_excel(text, file.filename)
    
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="converted.xlsx"
    )
