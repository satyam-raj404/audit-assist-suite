import uuid
from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import UserLog
from app.schemas import UserLogCreate, UserLogResponse
from typing import List

router = APIRouter(prefix="/api/logs", tags=["user_logs"])


def log_activity(db: Session, user_id: str | None, action: str, module: str, run_id: str | None = None, details: str | None = None):
    """Utility function for other routers to log user activity."""
    entry = UserLog(
        user_id=uuid.UUID(user_id) if user_id else None,
        action=action,
        module=module,
        run_id=uuid.UUID(run_id) if run_id else None,
        details=details,
    )
    db.add(entry)
    db.commit()


@router.post("", response_model=UserLogResponse)
async def create_user_log(data: UserLogCreate, db: Session = Depends(get_db)):
    entry = UserLog(
        user_id=uuid.UUID(data.user_id) if data.user_id else None,
        action=data.action,
        module=data.module,
        run_id=uuid.UUID(data.run_id) if data.run_id else None,
        details=data.details,
    )
    db.add(entry)
    db.commit()
    db.refresh(entry)
    return entry


@router.get("/{user_id}", response_model=List[UserLogResponse])
async def get_user_logs(user_id: uuid.UUID, db: Session = Depends(get_db)):
    return db.query(UserLog).filter(UserLog.user_id == user_id).order_by(UserLog.created_at.desc()).all()
