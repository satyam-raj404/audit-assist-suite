import uuid
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import UserProfile
from app.schemas import UserProfileCreate, UserProfileResponse

router = APIRouter(prefix="/api/users", tags=["user_onboarding"])


@router.post("/profile", response_model=UserProfileResponse)
async def save_user_profile(data: UserProfileCreate, db: Session = Depends(get_db)):
    uid = uuid.UUID(data.user_id)

    # Upsert: update if exists, create if not
    existing = db.query(UserProfile).filter(UserProfile.user_id == uid).first()
    if existing:
        existing.team_name = data.team_name
        existing.partner_name = data.partner_name
        existing.client_name = data.client_name
        existing.sector = data.sector
        existing.employee_id = data.employee_id
        db.commit()
        db.refresh(existing)
        return existing

    profile = UserProfile(
        user_id=uid,
        team_name=data.team_name,
        partner_name=data.partner_name,
        client_name=data.client_name,
        sector=data.sector,
        employee_id=data.employee_id,
    )
    db.add(profile)
    db.commit()
    db.refresh(profile)
    return profile


@router.get("/profile/{user_id}", response_model=UserProfileResponse)
async def get_user_profile(user_id: uuid.UUID, db: Session = Depends(get_db)):
    profile = db.query(UserProfile).filter(UserProfile.user_id == user_id).first()
    if not profile:
        raise HTTPException(status_code=404, detail="Profile not found")
    return profile
