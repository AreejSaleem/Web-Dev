from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
import shutil
import os
import re
from win32com.client import Dispatch
#from win32 import client
import zipfile
from io import BytesIO
import io
from contextlib import asynccontextmanager
from ppt_to_word import ppt_to_word
from extractedimages import (
    extract_images_from_ppt,
    extract_images_from_word,
    extract_images_from_pdf,
)
from compressorv1 import save_temp_file, create_zip
import openai
from fastapi import FastAPI, Form, HTTPException
from pptx import Presentation
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
#the import unicorn was commented if any issue arises comment that again
#import uvicorn
from presentation_generator import create_presentation, remove_slides, save_presentation
from presentation_generator import extract_text_from_word
from presentation_generator import extract_text_from_pdf
from presentation_generator import generate_title
import tempfile
from pathlib import Path
from convert import convert_docx_to_pdf, convert_pdf_to_docx
import logging
from google.oauth2 import id_token
from google.auth.transport import requests
GOOGLE_CLIENT_ID = "YOUR_GOOGLE_CLIENT_ID"

app = FastAPI()

#inlcuding loggings 
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

@asynccontextmanager
async def lifespan(app: FastAPI):
    # This will run on startup and shutdown
    yield
    # Clean up the extracted images when the app shuts down
    shutil.rmtree("extracted_images")
    os.makedirs("extracted_images", exist_ok=True)

app = FastAPI(lifespan=lifespan)

############### Compressor ##################################

@app.post("/compress_ppt/")
async def compress_ppt(file: UploadFile = File(...)):
    temp_file_path = save_temp_file(file)
    zip_file_path = create_zip(temp_file_path, file.filename)
    os.remove(temp_file_path)
    return FileResponse(zip_file_path, media_type="application/zip", filename=zip_file_path)

#####################################################################

############### Extracted Images ##################################

EXTRACTION_PATH = "extracted_images"
os.makedirs(EXTRACTION_PATH, exist_ok=True)

@app.post("/extract-images/")
async def extract_images(file_type: str = Form(...), file: UploadFile = File(...)):
    if file_type not in ["ppt", "word", "pdf"]:
        raise HTTPException(status_code=400, detail="Unsupported file type. Choose 'ppt', 'word', or 'pdf'.")
    
    temp_file_path = os.path.join(EXTRACTION_PATH, file.filename)
    with open(temp_file_path, "wb") as temp_file:
        shutil.copyfileobj(file.file, temp_file)

    if file_type == "ppt":
        image_paths = extract_images_from_ppt(temp_file_path)
    elif file_type == "word":
        image_paths = extract_images_from_word(temp_file_path)
    else:
        image_paths = extract_images_from_pdf(temp_file_path)

    os.remove(temp_file_path)
    return {"image_paths": image_paths}


@app.get("/download-all-images/")
async def download_all_images():
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for root, _, files in os.walk(EXTRACTION_PATH):
            for file_name in files:
                file_path = os.path.join(root, file_name)
                zip_file.write(file_path, arcname=file_name)

    zip_buffer.seek(0)

    # Set appropriate headers to trigger download
    headers = {
        "Content-Disposition": "attachment; filename=extracted_images.zip",
        "Content-Type": "application/zip"
    }
    return StreamingResponse(zip_buffer, headers=headers)
###########################################################

######################## PPT to Word #######################

@app.post("/convert/")
async def convert_ppt_to_word(file: UploadFile = File(...)):
    ppt_file = await file.read()
    word_doc = ppt_to_word(io.BytesIO(ppt_file))
    
    output_path = "output_document.docx"
    word_doc.save(output_path)
    
    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename="output_document.docx"
    )

#########################################################################

################### Presentation Generator ##############################

# Set up CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:8000",
        "http://127.0.0.1:5500",
        "http://127.0.0.1:8000",
        "https://slidestream.site",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



@app.post("/generate_presentation/")
async def generate_presentation_endpoint(
    topic: str = Form(...),
    slide_count: int = Form(...),
    template_choice: str = Form(...)
):
    """AI Presentation Generator.
        Choose a template:
    Minimalist,
    Corporate,
    Creative,
    Innovative,
    Aster,
    Business1,
    Business2,
    Business3,
    Business4,
    Business5,
    Student1,
    Student2,
    Student3,
    Student4,
    Student5,
    Corporate1,
    Corporate2,
    Corporate3,
    Corporate4,
    Corporate5,
    Creative1,
    Creative2,
    Creative3,
    Creative4,
    Creative5
"""
    if slide_count < 5 or slide_count > 10:
        raise HTTPException(status_code=400, detail="Slide count must be between 5 and 18.")

    try:
        prs = create_presentation(topic, slide_count, template_choice)
        remove_slides(prs, [0, 3])
        
        # Save presentation file in a known directory (e.g., 'generated_presentations')
        save_dir = 'generated_presentations'
        os.makedirs(save_dir, exist_ok=True)
        filename = f"{topic.replace(' ', '_')}.pptx"
        save_path = os.path.join(save_dir, filename)
        save_presentation(prs, save_path)

        # Ensure the file can be served correctly
        return FileResponse(
            path=save_path,
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
#####################################
#endpoint for Docs to PPT
@app.post("/generate_presentation_from_docs/")
async def generate_presentation_from_docs_endpoint(file: UploadFile = File(...), template_choice: str = Form(...), slide_count: int = Form(10)):
    """AI Presentation Generator.
        Choose a template:
    Minimalist,
    Corporate,
    Creative,
    Innovative,
    Aster,
    Business1,
    Business2,
    Business3,
    Business4,
    Business5,
    Student1,
    Student2,
    Student3,
    Student4,
    Student5,
    Corporate1,
    Corporate2,
    Corporate3,
    Corporate4,
    Corporate5,
    Creative1,
    Creative2,
    Creative3,
    Creative4,
    Creative5
    """
    
    # Log the incoming request details
    logger.info(f"Received request with file: {file.filename}, template: {template_choice}, slide_count: {slide_count}")

    # Validate slide_count
    if slide_count < 5 or slide_count > 18:
        logger.error("Invalid slide count, it must be between 5 and 18")
        raise HTTPException(status_code=400, detail="Slide count must be between 5 and 18.")

    try:
        # Save the uploaded file
        save_dir = "generated_presentations"
        os.makedirs(save_dir, exist_ok=True)
        file_path = os.path.join(save_dir, file.filename)
        with open(file_path, "wb") as f:
            f.write(await file.read())
        logger.info(f"File saved to {file_path}")

        # Extract content from the uploaded file
        if file_path.endswith('.pdf'):
            file_content = extract_text_from_pdf(file_path)
            logger.info("Extracted text from PDF.")
        elif file_path.endswith('.docx'):
            file_content = extract_text_from_word(file_path)
            logger.info("Extracted text from DOCX.")
        else:
            logger.error("Unsupported file format.")
            raise HTTPException(status_code=400, detail="Unsupported file format. Please upload a PDF or DOCX file.")

        # Generate the title based on the content
        title = generate_title(file_content)
        logger.info(f"Generated title: {title}")

        # Sanitize the title to ensure it's valid for a file path
        sanitized_title = re.sub(r'[<>:"/\\|?*]', '_', title)  # Replace invalid characters with underscores

        # Create the presentation
        prs = create_presentation(sanitized_title, slide_count, template_choice)
        logger.info("Created presentation object.")

        # Process the extracted content and add slides
        slide_texts = file_content.split('\n')  # Split the content into lines for slides
        for i in range(min(slide_count, len(slide_texts))):
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Using layout without placeholders (index 5)
            text_box = slide.shapes.add_textbox(left=100, top=100, width=500, height=200)
            text_frame = text_box.text_frame
            text_frame.text = slide_texts[i]
            logger.info(f"Added slide {i+1} with content: {slide_texts[i]}")

        # Save the presentation with a sanitized filename
        presentation_filename = f"{sanitized_title}.pptx"
        save_path = os.path.join(save_dir, presentation_filename)
        save_presentation(prs, save_path)
        logger.info(f"Presentation saved to {save_path}")

        # Serve the presentation file as a download
        return FileResponse(
            path=save_path,
            headers={"Content-Disposition": f'attachment; filename="{presentation_filename}"'},
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
        raise HTTPException(status_code=500, detail=f"An error occurred: {str(e)}")


#####End of new endpoint ##########
# Endpoint to convert DOCX to PDF
@app.post("/convert-docx-to-pdf/")
async def convert_docx_to_pdf_endpoint(file: UploadFile = File(...)):
    # Get the user's Downloads folder path
    #downloads_folder = str(Path.home() / "Downloads")
    #downloads_folder = Path(r"C:\Users\Public\Public Downloads")
    downloads_folder = "C:/Users/Public/Public Downloads" 

    # Create temporary files for input
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_input:
        tmp_input_path = tmp_input.name
        with open(tmp_input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    
    # Define the output PDF file path in the Downloads folder
    output_file = os.path.join(downloads_folder, f"{os.path.splitext(file.filename)[0]}.pdf")
    
    # Convert the DOCX to PDF
    pdf_path = convert_docx_to_pdf(tmp_input_path, output_file)
    
    if pdf_path:
        # Return the converted PDF file to the user
        return FileResponse(pdf_path, media_type='application/pdf', filename=os.path.basename(pdf_path))
    else:
        return {"error": "File conversion failed."}

# Endpoint to convert PDF to DOCX
@app.post("/convert-pdf-to-docx/")
async def convert_pdf_to_docx_endpoint(file: UploadFile = File(...)):
    # Get the user's Downloads folder path
    downloads_folder = str(Path.home() / "Downloads")

    # Save the uploaded PDF file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_input:
        tmp_input_path = tmp_input.name
        with open(tmp_input_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    
    # Define the output DOCX file path in the Downloads folder
    output_file = os.path.join(downloads_folder, f"{os.path.splitext(file.filename)[0]}.docx")

    # Convert the PDF to DOCX using pdf2docx
    docx_path = convert_pdf_to_docx(tmp_input_path, output_file)

    if docx_path:
        # Return the converted DOCX file to the user
        return FileResponse(docx_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=os.path.basename(docx_path))
    else:
        return {"error": "PDF to DOCX conversion failed."}
    ####checking word access from win32com.client import Dispatch

    
@app.get("/")
def root():
    return {"message": "Welcome to the Presentation Generator API"}
 
####################################################################################################################
######################## Converter Word/PDF ########################################################





if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="localhost", port=8000)
