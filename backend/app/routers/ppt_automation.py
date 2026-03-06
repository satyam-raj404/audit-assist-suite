import uuid
import sys
from pathlib import Path
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import AuditRun, UploadedFile
from app.schemas import PPTAutomationRequest, PPTAutomationResponse, AuditStepStatus

router = APIRouter(prefix="/api/ppt", tags=["ppt_automation"])

# Absolute path to the automation scripts root (contains IA_Report, Concurrent_Audit_Report, etc.)
SCRIPTS_ROOT = Path(__file__).resolve().parents[3] / "Report Utility Automation Scripts"

PPT_STEPS = [
    {"id": "parse",        "label": "Parsing tracker file"},
    {"id": "validate",     "label": "Validating input data"},
    {"id": "load_template","label": "Loading PPT template"},
    {"id": "populate",     "label": "Populating template with data"},
    {"id": "format",       "label": "Applying formatting rules"},
    {"id": "save",         "label": "Saving output file"},
]


def _get_file_path(db, file_id_str: str) -> str:
    """Look up a file's disk path by its UUID string."""
    db_file = db.query(UploadedFile).filter(UploadedFile.id == uuid.UUID(file_id_str)).first()
    if not db_file:
        raise FileNotFoundError(f"Uploaded file {file_id_str} not found in database")
    return db_file.path


def _ensure_scripts_on_path():
    scripts_str = str(SCRIPTS_ROOT)
    if scripts_str not in sys.path:
        sys.path.insert(0, scripts_str)


def run_ppt_task(run_id: uuid.UUID, db_url: str, pptx_override_path: str | None = None):
    """Background task that calls the real automation scripts."""
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
        run.progress = 10
        db.commit()

        # Resolve file paths from the DB
        excel_path = _get_file_path(db, str(run.excel_file_id))
        if pptx_override_path:
            pptx_path = pptx_override_path
        elif run.pptx_file_id:
            pptx_path = _get_file_path(db, str(run.pptx_file_id))
        else:
            pptx_path = None

        _ensure_scripts_on_path()

        audit_type = run.audit_type
        utility_type = run.utility_type or ""
        report_type = run.report_type or "Both"
        output_path = run.output_path
        month = run.month or ""
        year = run.year or ""

        run.progress = 25
        db.commit()

        # Dispatch to the correct module — mirrors FS_Utility.main() without importing it
        # (FS_Utility.py has module-level Tkinter code that would crash on import)
        result_path = None

        if audit_type == "Internal Audit" and utility_type == "Report":
            from IA_Report.IA_Report_Utility_v2 import IA_Report_Utility_Main
            result_path = IA_Report_Utility_Main(excel_path, pptx_path, output_path)

        elif audit_type == "Internal Audit-Zensar" and utility_type == "Report":
            from IA_Report_Template2.inspect_placeholders import IA_Report_Utility_MainT2
            result_path = IA_Report_Utility_MainT2(excel_path, pptx_path, output_path)

        elif audit_type == "Concurrent Audit" and utility_type == "Report":
            from Concurrent_Audit_Report.Concurrent_Report import Concurrent_Report_Main
            result_path = Concurrent_Report_Main(excel_path, pptx_path, report_type, output_path, month, year)

        elif audit_type == "Concurrent Audit" and utility_type == "Dashboard":
            from Concurrent_Audit_Dashboard.Concurrent_Dashboard import Concurrent_Dashboard_Main
            result_path = Concurrent_Dashboard_Main(excel_path, pptx_path, output_path)

        elif audit_type == "ICOFR" and utility_type == "Dashboard":
            from ICOFR import main as ICOFR_main
            result_path = ICOFR_main(excel_path, pptx_path, output_path)

        else:
            raise ValueError(f"Unsupported combination: audit_type='{audit_type}', utility_type='{utility_type}'")

        run.progress = 100
        run.status = "complete"
        run.report_path = result_path if result_path else output_path
        db.commit()

        from app.routers.user_logs import log_activity
        log_activity(db, str(run.user_id) if run.user_id else None, "ppt_complete", "ppt_automation", str(run.id), f"PPT generation completed: {audit_type}")

    except Exception as e:
        run = db.query(AuditRun).filter(AuditRun.id == run_id).first()
        if run:
            run.status = "error"
            run.report_path = str(e)
            db.commit()
        from app.routers.error_logging import log_error
        log_error(db, None, "ppt_automation", str(e), "")
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
        utility_type=data.utility_type,
        report_type=data.report_type,
        excel_file_id=uuid.UUID(data.excel_file_id),
        pptx_file_id=uuid.UUID(data.pptx_file_id) if data.pptx_file_id else None,
        month=data.month,
        year=data.year,
        output_path=data.output_path,
        user_id=uuid.UUID(data.user_id) if data.user_id else None,
    )
    db.add(run)
    db.commit()
    db.refresh(run)

    from app.routers.user_logs import log_activity
    log_activity(db, data.user_id, "ppt_start", "ppt_automation", str(run.id), f"Started: {data.audit_type}")

    from app.database import DATABASE_URL
    background_tasks.add_task(run_ppt_task, run.id, DATABASE_URL, data.pptx_path or None)

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
        id=run.id, status=run.status, progress=run.progress,
        steps=steps, report_path=run.report_path
    )


@router.post("/{run_id}/cancel")
async def cancel_ppt(run_id: uuid.UUID, db: Session = Depends(get_db)):
    run = db.query(AuditRun).filter(AuditRun.id == run_id).first()
    if not run:
        raise HTTPException(status_code=404, detail="Run not found")
    run.status = "error"
    db.commit()
    return {"detail": "Run cancelled"}
