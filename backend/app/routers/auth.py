import uuid
from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from sqlalchemy import text
from app.database import get_db
from app.schemas import RegisterRequest, RegisterResponse, LoginRequest, LoginResponse
import bcrypt

router = APIRouter(prefix="/api/auth", tags=["auth"])


def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()


def verify_password(plain: str, hashed: str) -> bool:
    return bcrypt.checkpw(plain.encode(), hashed.encode())


@router.post("/register", response_model=RegisterResponse, status_code=status.HTTP_201_CREATED)
async def register(data: RegisterRequest, db: Session = Depends(get_db)):
    # Check if email already exists using raw SQL
    result = db.execute(
        text("SELECT id FROM users WHERE email = :email"),
        {"email": data.email},
    ).fetchone()

    if result:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="An account with this email already exists.",
        )

    new_id = uuid.uuid4()
    hashed = hash_password(data.password)

    db.execute(
        text(
            "INSERT INTO users (id, name, email, hashed_password, created_at) "
            "VALUES (:id, :name, :email, :hashed_password, NOW())"
        ),
        {"id": new_id, "name": data.name, "email": data.email, "hashed_password": hashed},
    )
    db.commit()

    return RegisterResponse(id=new_id, name=data.name, email=data.email)


@router.post("/login", response_model=LoginResponse)
async def login(data: LoginRequest, db: Session = Depends(get_db)):
    row = db.execute(
        text("SELECT id, name, email, hashed_password FROM users WHERE email = :email"),
        {"email": data.email},
    ).fetchone()

    if not row or not verify_password(data.password, row.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid email or password.",
        )

    return LoginResponse(id=row.id, name=row.name, email=row.email)
