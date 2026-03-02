import os
import uuid
import shutil
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import TemplateRequest
from app.schemas import TemplateRequestResponse
from typing import List, Optional

router = APIRouter(prefix="/api/templates", tags=["template_requests"])

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "..", "..", "uploads", "templates")
os.makedirs(UPLOAD_DIR, exist_ok=True)


@router.post("/request", response_model=TemplateRequestResponse)
async def submit_template_request(
    file: UploadFile = File(...),
    spoc_email: str = Form(...),
    team_wide: bool = Form(False),
    user_id: Optional[str] = Form(None),
    db: Session = Depends(get_db),
):
    # Save file to disk
    file_id = str(uuid.uuid4())
    file_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    record = TemplateRequest(
        file_name=file.filename,
        file_path=file_path,
        spoc_email=spoc_email,
        team_wide=team_wide,
        user_id=uuid.UUID(user_id) if user_id else None,
    )
    db.add(record)
    db.commit()
    db.refresh(record)

    return record


@router.get("/requests", response_model=List[TemplateRequestResponse])
async def list_template_requests(db: Session = Depends(get_db)):
    return db.query(TemplateRequest).order_by(TemplateRequest.created_at.desc()).all()
