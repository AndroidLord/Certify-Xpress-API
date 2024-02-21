from fastapi import FastAPI, Query, status, UploadFile, File 
import config, s3config, schema, os, openpyxl
from util import replace_text_in_docx, convert_to_pdf
from datetime import datetime
from fastapi.responses import JSONResponse
from routers import certification

app = FastAPI()


@app.get("/")
async def root():
    return {"message": "This is the root of the CertifyXpresAPI, please use the /docs endpoint to see the API documentation."}


app.include_router(certification.router, prefix="/api/v1")

