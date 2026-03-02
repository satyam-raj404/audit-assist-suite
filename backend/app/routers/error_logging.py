import uuid
from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import ErrorLog
from app.schemas import ErrorLogCreate, ErrorLogResponse
from typing import List

router = APIRouter(prefix="/api/errors", tags=["error_logging"])


def log_error(db: Session, user_id: str | None, module: str, message: str, trace: str | None = None):
    """Utility function for other routers to log errors."""
    entry = ErrorLog(
        user_id=uuid.UUID(user_id) if user_id else None,
        module=module,
        error_message=message,
        stack_trace=trace,
    )
    db.add(entry)
    db.commit()


@router.post("/log", response_model=ErrorLogResponse)
async def create_error_log(data: ErrorLogCreate, db: Session = Depends(get_db)):
    entry = ErrorLog(
        user_id=uuid.UUID(data.user_id) if data.user_id else None,
        module=data.module,
        error_message=data.error_message,
        stack_trace=data.stack_trace,
    )
    db.add(entry)
    db.commit()
    db.refresh(entry)
    return entry


@router.get("", response_model=List[ErrorLogResponse])
async def list_errors(limit: int = 50, db: Session = Depends(get_db)):
    return db.query(ErrorLog).order_by(ErrorLog.created_at.desc()).limit(limit).all()
