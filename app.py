import os
import shutil
import uuid
from fastapi import FastAPI, File, UploadFile, Request, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from index import process_resume

app = FastAPI()

# Directories
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
TEMPLATE_DOCX = "ravan.docx"  # Ensure this matches your template filename

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory="static"), name="static")

# Templates
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# In-memory storage for task status
task_status = {}

def process_resume_task(task_id: str, pdf_path: str, template_path: str, output_path: str):
    try:
        def update_status(msg):
            task_status[task_id]["status"] = msg

        task_status[task_id]["status"] = "Starting..."
        process_resume(pdf_path, template_path, output_path, status_callback=update_status)
        
        task_status[task_id]["status"] = "Completed"
        task_status[task_id]["download_url"] = f"/download/{os.path.basename(output_path)}"
        
    except Exception as e:
        task_status[task_id]["status"] = f"Error: {str(e)}"
        task_status[task_id]["error"] = str(e)

@app.post("/convert")
async def convert_resume(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Generate unique IDs for filenames
    unique_id = str(uuid.uuid4())
    pdf_filename = f"{unique_id}_{file.filename}"
    pdf_path = os.path.join(UPLOAD_DIR, pdf_filename)
    
    # Save uploaded file
    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    # Define output path
    output_filename = f"Generated_Resume_{unique_id}.docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    
    # Initialize task status
    task_status[unique_id] = {"status": "Queued"}
    
    # Start background task
    background_tasks.add_task(process_resume_task, unique_id, pdf_path, TEMPLATE_DOCX, output_path)
    
    return {"task_id": unique_id}

@app.get("/status/{task_id}")
async def get_status(task_id: str):
    return task_status.get(task_id, {"status": "Unknown Task"})


@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return FileResponse(file_path, filename=filename, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    return {"error": "File not found"}
