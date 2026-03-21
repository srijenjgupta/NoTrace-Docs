from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import tempfile
import os
import platform
import subprocess
import zipfile
import fitz  # PyMuPDF
from pdf2docx import Converter
import pdfplumber
import pandas as pd
from PIL import Image
import openpyxl
import json

app = FastAPI(title="Zero-Retention PDF API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Original-Size", "X-New-Size"]
)

def cleanup_files(file_paths):
    for path in file_paths:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception as e:
                print(f"Cleanup note: Could not remove {path} immediately. {e}")

@app.post("/organize-pdf")
async def organize_pdf(files: List[UploadFile] = File(...), order: str = Form(...)):
    """The Smart Engine: Handles Merging, Splitting, Reordering, and Deleting simultaneously."""
    docs = []
    temp_merged_path = None
    try:
        # 1. Parse the order instructions from the frontend
        order_list = json.loads(order)
        merged_pdf = fitz.open()
        
        # 2. Open all uploaded PDFs into RAM
        for f in files:
            docs.append(fitz.open("pdf", await f.read()))

        # 3. Build the new PDF based on the exact user-dragged order
        for item in order_list:
            file_idx = int(item['file_idx'])
            page_idx = int(item['page'])
            doc = docs[file_idx]
            # Insert just that specific page
            merged_pdf.insert_pdf(doc, from_page=page_idx, to_page=page_idx)

        # 4. Save and cleanup
        temp_merged_path = tempfile.mktemp(suffix=".pdf")
        merged_pdf.save(temp_merged_path)
        
        for doc in docs:
            doc.close()
        merged_pdf.close()

        return FileResponse(temp_merged_path, filename="organized_custom.pdf", 
                            background=lambda: cleanup_files([temp_merged_path]))
    except Exception as e:
        for doc in docs: doc.close()
        if temp_merged_path: cleanup_files([temp_merged_path])
        raise HTTPException(status_code=500, detail=f"Organize failed: {str(e)}")

@app.post("/convert-to-word")
async def convert_to_word(file: UploadFile = File(...)):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(await file.read())
            temp_pdf_path = temp_pdf.name
        temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()
        return FileResponse(temp_docx_path, filename="converted.docx", background=lambda: cleanup_files([temp_pdf_path, temp_docx_path]))
    except Exception as e:
        if 'temp_pdf_path' in locals(): cleanup_files([temp_pdf_path])
        raise HTTPException(status_code=500, detail=f"Word conversion failed: {str(e)}")

@app.post("/convert-to-excel")
async def convert_to_excel(file: UploadFile = File(...)):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(await file.read())
            temp_pdf_path = temp_pdf.name
        temp_excel_path = temp_pdf_path.replace(".pdf", ".xlsx")
        all_tables = []
        with pdfplumber.open(temp_pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if table: 
                        df = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df)
        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df.to_excel(temp_excel_path, index=False, engine='openpyxl')
        else:
            pd.DataFrame([{"Message": "No tables found"}]).to_excel(temp_excel_path, index=False, engine='openpyxl')
        return FileResponse(temp_excel_path, filename="extracted_tables.xlsx", background=lambda: cleanup_files([temp_pdf_path, temp_excel_path]))
    except Exception as e:
        if 'temp_pdf_path' in locals(): cleanup_files([temp_pdf_path])
        raise HTTPException(status_code=500, detail=f"Excel conversion failed: {str(e)}")

@app.post("/convert-to-pdf")
async def convert_office_to_pdf(file: UploadFile = File(...)):
    try:
        ext = os.path.splitext(file.filename)[1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as temp_input:
            temp_input.write(await file.read())
            temp_input_path = temp_input.name
        output_dir = os.path.dirname(temp_input_path)
        temp_pdf_path = temp_input_path.replace(ext, ".pdf")
        libre_path = r"C:\Program Files\LibreOffice\program\soffice.exe" if platform.system() == "Windows" else "libreoffice"
        subprocess.run([libre_path, "--headless", "--convert-to", "pdf", temp_input_path, "--outdir", output_dir], check=True)
        return FileResponse(temp_pdf_path, filename=f"converted.pdf", background=lambda: cleanup_files([temp_input_path, temp_pdf_path]))
    except Exception as e:
        if 'temp_input_path' in locals(): cleanup_files([temp_input_path])
        raise HTTPException(status_code=500, detail=f"Office conversion failed: {str(e)}")

@app.post("/img-to-pdf")
async def convert_img_to_pdf(files: List[UploadFile] = File(...)):
    try:
        image_list = [Image.open(f.file).convert('RGB') for f in files]
        temp_pdf_path = tempfile.mktemp(suffix=".pdf")
        if image_list:
            image_list[0].save(temp_pdf_path, save_all=True, append_images=image_list[1:])
        return FileResponse(temp_pdf_path, filename="images.pdf", background=lambda: cleanup_files([temp_pdf_path]))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Image to PDF failed: {str(e)}")

@app.post("/pdf-to-img")
async def convert_pdf_to_img(file: UploadFile = File(...)):
    temp_pdf_path = tempfile.mktemp(suffix=".pdf")
    temp_zip_path = tempfile.mktemp(suffix=".zip")
    try:
        with open(temp_pdf_path, "wb") as f: f.write(await file.read())
        doc = fitz.open(temp_pdf_path)
        with zipfile.ZipFile(temp_zip_path, 'w') as zipf:
            for i in range(len(doc)):
                pix = doc[i].get_pixmap(dpi=150)
                img_path = tempfile.mktemp(suffix=".png")
                pix.save(img_path)
                zipf.write(img_path, f"page_{i+1}.png")
                os.remove(img_path)
        doc.close()
        return FileResponse(temp_zip_path, filename="images.zip", background=lambda: cleanup_files([temp_pdf_path, temp_zip_path]))
    except Exception as e:
        cleanup_files([temp_pdf_path, temp_zip_path])
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/compress")
async def compress_file(file: UploadFile = File(...), level: str = Form("recommended")):
    try:
        ext = os.path.splitext(file.filename)[1].lower()
        file_bytes = await file.read()
        orig_size = len(file_bytes)
        temp_in = tempfile.mktemp(suffix=ext)
        with open(temp_in, "wb") as f: f.write(file_bytes)
        temp_out = tempfile.mktemp(suffix=ext)
        
        if ext == ".pdf":
            doc = fitz.open(temp_in)
            doc.save(temp_out, deflate=True, garbage=4 if level=="extreme" else 2)
            doc.close()
        elif ext in [".jpg", ".jpeg", ".png"]:
            img = Image.open(temp_in).convert("RGB")
            temp_out = temp_out.replace(".png", ".jpg")
            img.save(temp_out, "JPEG", optimize=True, quality=30 if level=="extreme" else 65)
        
        headers = {"X-Original-Size": str(orig_size), "X-New-Size": str(os.path.getsize(temp_out))}
        return FileResponse(temp_out, filename=f"compressed_{file.filename}", headers=headers, background=lambda: cleanup_files([temp_in, temp_out]))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/unlock-excel")
async def unlock_excel(file: UploadFile = File(...)):
    temp_input_path = None
    temp_output_path = None
    wb = None
    try:
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in [".xlsx", ".xlsm"]: raise HTTPException(status_code=400, detail="Only .xlsx/.xlsm allowed.")
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as temp_input:
            temp_input.write(await file.read())
            temp_input_path = temp_input.name
        temp_output_path = temp_input_path.replace(ext, f"_unlocked{ext}")
        wb = openpyxl.load_workbook(temp_input_path)
        wb.security = None
        for sheet_name in wb.sheetnames:
            wb[sheet_name].protection.disable()
        wb.save(temp_output_path)
        return FileResponse(temp_output_path, filename=f"unlocked_{file.filename}", background=lambda: cleanup_files([temp_input_path, temp_output_path]))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unlock failed: {str(e)}")
    finally:
        if wb: wb.close()