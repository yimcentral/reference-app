from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import uvicorn
import os
import pandas as pd
from reference_utils import generate_reference_docx

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate", response_class=FileResponse)
async def generate(
    request: Request,
    prefix: str = Form(...),
    agency: str = Form(...),
    proceeding: str = Form(...),
):
    df = pd.read_csv("sample_docket.csv")  # replace with database pull later
    doc_path = generate_reference_docx(df, prefix, agency, proceeding)
    return FileResponse(path=doc_path, filename="Reference_List.docx")

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000)
