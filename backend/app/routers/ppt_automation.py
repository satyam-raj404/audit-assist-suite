import uuid
import time
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import AuditRun
from app.schemas import PPTAutomationRequest, PPTAutomationResponse, AuditStepStatus

router = APIRouter(prefix="/api/ppt", tags=["ppt_automation"])

PPT_STEPS = [
    {"id": "parse", "label": "Parsing tracker file"},
    {"id": "validate", "label": "Validating input data"},
    {"id": "load_template", "label": "Loading PPT template"},
    {"id": "populate", "label": "Populating template with data"},
    {"id": "format", "label": "Applying formatting rules"},
    {"id": "save", "label": "Saving output file"},
]


def run_ppt_task(run_id: uuid.UUID, db_url: str):
    """Background stub that simulates the 6 processing steps."""
    from sqlalchemy import create_engine
    from sqlalchemy.orm import sessionmaker

    engine = create_engine(db_url)
    Sess = sessionmaker(bind=engine)
    db = Sess()

    try:
        run = db.query(AuditRun).filter(AuditRun.id == run_id).first()
        if not run:
            return

        run.status = "running"
        db.commit()

        for i in range(len(PPT_STEPS)):
            time.sleep(1)
            run.progress = ((i + 1) / len(PPT_STEPS)) * 100
            db.commit()

        run.status = "complete"
        run.progress = 100
        run.report_path = f"{run.output_path}/report_{run.id}.pptx"
        db.commit()

        # Log activity
        from app.routers.user_logs import log_activity
        log_activity(db, str(run.user_id) if run.user_id else None, "ppt_complete", "ppt_automation", str(run.id), "PPT generation completed")
    except Exception:
        run.status = "error"
        db.commit()
    finally:
        db.close()


@router.post("/start", response_model=PPTAutomationResponse)
async def start_ppt(
    data: PPTAutomationRequest,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
):
    run = AuditRun(
        audit_type=data.audit_type,
        sub_audit_sector=data.sub_audit_sector,
        template_id=data.template_id,
        output_path=data.output_path,
        user_id=uuid.UUID(data.user_id) if data.user_id else None,
    )
    db.add(run)
    db.commit()
    db.refresh(run)

    # Log activity
    from app.routers.user_logs import log_activity
    log_activity(db, data.user_id, "ppt_start", "ppt_automation", str(run.id), f"Started PPT automation: {data.audit_type}")

    from app.database import DATABASE_URL
    background_tasks.add_task(run_ppt_task, run.id, DATABASE_URL)

    steps = [AuditStepStatus(id=s["id"], label=s["label"], status="pending") for s in PPT_STEPS]
    return PPTAutomationResponse(id=run.id, status=run.status, progress=run.progress, steps=steps)


@router.get("/{run_id}/status", response_model=PPTAutomationResponse)
async def get_ppt_status(run_id: uuid.UUID, db: Session = Depends(get_db)):
    run = db.query(AuditRun).filter(AuditRun.id == run_id).first()
    if not run:
        raise HTTPException(status_code=404, detail="Run not found")

    current_step = int((run.progress / 100) * len(PPT_STEPS))
    steps = []
    for i, s in enumerate(PPT_STEPS):
        if i < current_step:
            status = "complete"
        elif i == current_step and run.status == "running":
            status = "running"
        else:
            status = "pending"
        steps.append(AuditStepStatus(id=s["id"], label=s["label"], status=status))

    return PPTAutomationResponse(
        id=run.id, status=run.status, progress=run.progress, steps=steps, report_path=run.report_path
    )


@router.post("/{run_id}/cancel")
async def cancel_ppt(run_id: uuid.UUID, db: Session = Depends(get_db)):
    run = db.query(AuditRun).filter(AuditRun.id == run_id).first()
    if not run:
        raise HTTPException(status_code=404, detail="Run not found")

    run.status = "error"
    db.commit()
    return {"detail": "Run cancelled"}
