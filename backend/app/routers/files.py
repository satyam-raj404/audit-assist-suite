import os
import uuid
import asyncio
from fastapi import APIRouter, UploadFile, File, Depends, HTTPException
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import UploadedFile
from app.schemas import FileResponse

router = APIRouter(prefix="/api/files", tags=["files"])

UPLOAD_DIR = os.getenv("UPLOAD_DIR", "./uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


@router.post("/upload", response_model=FileResponse)
async def upload_file(file: UploadFile = File(...), db: Session = Depends(get_db)):
    file_id = uuid.uuid4()
    file_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")

    content = await file.read()
    with open(file_path, "wb") as f:
        f.write(content)

    db_file = UploadedFile(
        id=file_id,
        name=file.filename,
        size=len(content),
        type=file.content_type or "application/octet-stream",
        path=file_path,
    )
    db.add(db_file)
    db.commit()
    db.refresh(db_file)

    return db_file


@router.get("/browse-folder")
async def browse_folder():
    """Open a native Windows folder picker dialog on the server machine and return the selected path."""
    def _pick() -> str:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(title="Select Output Folder")
        root.destroy()
        return path or ""

    loop = asyncio.get_running_loop()
    path = await loop.run_in_executor(None, _pick)
    return {"path": path}


@router.delete("/{file_id}")
async def delete_file(file_id: uuid.UUID, db: Session = Depends(get_db)):
    db_file = db.query(UploadedFile).filter(UploadedFile.id == file_id).first()
    if not db_file:
        raise HTTPException(status_code=404, detail="File not found")

    if os.path.exists(db_file.path):
        os.remove(db_file.path)

    db.delete(db_file)
    db.commit()
    return {"detail": "File deleted"}
