import os
from io import BytesIO
from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from processor import process_docx_bytes

APP_TOKEN = os.getenv('APP_TOKEN')
app = FastAPI(title="ADB HSU Formatter API", version="1.0.0")

# Allow your Vercel site to call this API (we can lock it later)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

def verify(x_token: str | None):
    if not APP_TOKEN:
        raise HTTPException(500, "Server not configured")
    if x_token != APP_TOKEN:
        raise HTTPException(401, "Unauthorized")

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/format")
async def format_docx(file: UploadFile = File(...), x_app_token: str | None = Header(None)):
    verify(x_app_token)
    data = await file.read()
    out = process_docx_bytes(data)
    fname = (file.filename or "document").rsplit(".", 1)[0] + ".hsu.docx"
    headers = {"Content-Disposition": f'attachment; filename="{fname}"'}
    return StreamingResponse(
        BytesIO(out),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers
    )
