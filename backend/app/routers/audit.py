import uuid
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import AuditRun
from app.schemas import AuditRequest, AuditResponse, AuditStepStatus
import time

router = APIRouter(prefix="/api/audit", tags=["audit"])

AUDIT_STEPS = [
    {"id": "validate", "label": "Validating input files"},
    {"id": "parse", "label": "Parsing document content"},
    {"id": "analyze", "label": "Running audit analysis"},
    {"id": "generate", "label": "Generating report"},
    {"id": "save", "label": "Saving to output folder"},
]


def run_audit_task(audit_id: uuid.UUID, db_url: str):
    """Background task that simulates audit processing.
    Replace this with your actual audit logic."""
    from sqlalchemy import create_engine
    from sqlalchemy.orm import sessionmaker

    engine = create_engine(db_url)
    Session = sessionmaker(bind=engine)
    db = Session()

    try:
        audit = db.query(AuditRun).filter(AuditRun.id == audit_id).first()
        if not audit:
            return

        audit.status = "running"
        db.commit()

        for i in range(5):
            time.sleep(1)  # Simulate work
            audit.progress = ((i + 1) / 5) * 100
            db.commit()

        audit.status = "complete"
        audit.progress = 100
        audit.report_path = f"{audit.output_path}/report_{audit.id}.pdf"
        db.commit()
    except Exception:
        audit.status = "error"
        db.commit()
    finally:
        db.close()


@router.post("/start", response_model=AuditResponse)
async def start_audit(
    data: AuditRequest,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
):
    audit = AuditRun(
        audit_type=data.audit_type,
        sub_audit_sector=data.sub_audit_sector,
        template_id=data.template_id,
        output_path=data.output_path,
    )
    db.add(audit)
    db.commit()
    db.refresh(audit)

    from app.database import DATABASE_URL
    background_tasks.add_task(run_audit_task, audit.id, DATABASE_URL)

    steps = [AuditStepStatus(id=s["id"], label=s["label"], status="pending") for s in AUDIT_STEPS]
    return AuditResponse(
        id=audit.id,
        status=audit.status,
        progress=audit.progress,
        steps=steps,
    )


@router.get("/{audit_id}/status", response_model=AuditResponse)
async def get_audit_status(audit_id: uuid.UUID, db: Session = Depends(get_db)):
    audit = db.query(AuditRun).filter(AuditRun.id == audit_id).first()
    if not audit:
        raise HTTPException(status_code=404, detail="Audit not found")

    current_step = int((audit.progress / 100) * 5)
    steps = []
    for i, s in enumerate(AUDIT_STEPS):
        if i < current_step:
            status = "complete"
        elif i == current_step and audit.status == "running":
            status = "running"
        else:
            status = "pending"
        steps.append(AuditStepStatus(id=s["id"], label=s["label"], status=status))

    return AuditResponse(
        id=audit.id,
        status=audit.status,
        progress=audit.progress,
        steps=steps,
        report_path=audit.report_path,
    )


@router.post("/{audit_id}/cancel")
async def cancel_audit(audit_id: uuid.UUID, db: Session = Depends(get_db)):
    audit = db.query(AuditRun).filter(AuditRun.id == audit_id).first()
    if not audit:
        raise HTTPException(status_code=404, detail="Audit not found")

    audit.status = "error"
    db.commit()
    return {"detail": "Audit cancelled"}
