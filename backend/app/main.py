from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.database import engine, Base
from app.routers import files, audit, reconciliation

# Create tables
Base.metadata.create_all(bind=engine)

app = FastAPI(title="KPMG Audit Tool API", version="1.0.0")

# CORS - allow React frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:8080", "http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(files.router)
app.include_router(audit.router)
app.include_router(reconciliation.router)


@app.get("/api/health")
async def health_check():
    return {"status": "healthy"}
