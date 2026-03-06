from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.database import engine, Base
from app.routers import files, ppt_automation, reconciliation, auth, template_req, user_onboarding, error_logging, user_logs

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
app.include_router(ppt_automation.router)
app.include_router(reconciliation.router)
app.include_router(auth.router)
app.include_router(template_req.router)
app.include_router(user_onboarding.router)
app.include_router(error_logging.router)
app.include_router(user_logs.router)


@app.get("/api/health")
async def health_check():
    return {"status": "healthy"}
