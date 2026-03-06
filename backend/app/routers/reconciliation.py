import uuid
import sys
from pathlib import Path
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse as FastAPIFileResponse
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import ReconciliationRun, UploadedFile
from app.schemas import ReconciliationRequest, ReconciliationResponse

router = APIRouter(prefix="/api/reconciliation", tags=["reconciliation"])

# Absolute path to the automation scripts root
SCRIPTS_ROOT = Path(__file__).resolve().parents[3] / "Report Utility Automation Scripts"


def _get_file_path(db, file_id_str: str) -> str:
    db_file = db.query(UploadedFile).filter(UploadedFile.id == uuid.UUID(file_id_str)).first()
    if not db_file:
        raise FileNotFoundError(f"Uploaded file {file_id_str} not found in database")
    return db_file.path


def _ensure_scripts_on_path():
    scripts_str = str(SCRIPTS_ROOT)
    if scripts_str not in sys.path:
        sys.path.insert(0, scripts_str)


def run_reconciliation_task(recon_id: uuid.UUID, db_url: str):
    """Background task that calls the real Excel analysis/reconciliation scripts."""
    from sqlalchemy import create_engine
    from sqlalchemy.orm import sessionmaker

    engine = create_engine(db_url)
    Sess = sessionmaker(bind=engine)
    db = Sess()

    try:
        recon = db.query(ReconciliationRun).filter(ReconciliationRun.id == recon_id).first()
        if not recon:
            return

        recon.status = "running"
        recon.progress = 10
        db.commit()

        source_path = _get_file_path(db, str(recon.source_file_id))
        target_path = _get_file_path(db, str(recon.target_file_id)) if recon.target_file_id else None
        output_path = recon.output_path
        audit_type = recon.audit_type

        _ensure_scripts_on_path()

        recon.progress = 25
        db.commit()

        result_path = None

        if audit_type == "FD - Sampling":
            from FD_Focused_Sampling.FD_Focused_Sampling import FD_Main
            result_path = FD_Main(source_path, output_path)

        elif audit_type == "KYC":
            from KYC.KYC import KYC_main
            result_path = KYC_main(source_path, output_path)

        elif audit_type == "CCIL":
            from CCIL.CCIL import CCIL_main
            result_path = CCIL_main(source_path, output_path)

        elif audit_type == "FOBO":
            from FOBO.FOBO import main as fo_bo_main
            result_path = fo_bo_main(source_path, output_path)

        elif audit_type == "IOA KPI1":
            from IOA_Sampling.KPI1.KPI1_script import main as kpi1_main
            result_path = kpi1_main(source_path, output_path)

        elif audit_type == "IOA KPI2":
            from IOA_Sampling.KPI2.KPI2_script import main as kpi2_main
            result_path = kpi2_main(source_path, output_path)

        elif audit_type == "IOA KPI3":
            from IOA_Sampling.KPI3.KPI3_script import main as kpi3_main
            result_path = kpi3_main(source_path, output_path)

        elif audit_type == "IOA KPI4":
            from IOA_Sampling.KPI4.KPI4_script import main as kpi4_main
            result_path = kpi4_main(source_path, output_path)

        elif audit_type == "Swift":
            from Swift.Reconv3 import recon_main
            result_path = recon_main(source_path, target_path, output_path)

        else:
            raise ValueError(f"Unsupported reconciliation audit_type: '{audit_type}'")

        recon.progress = 100
        recon.status = "complete"
        recon.report_path = result_path if result_path else output_path
        db.commit()

        from app.routers.user_logs import log_activity
        log_activity(db, str(recon.user_id) if recon.user_id else None, "reconciliation_complete", "reconciliation", str(recon.id), f"Completed: {audit_type}")

    except Exception as e:
        recon = db.query(ReconciliationRun).filter(ReconciliationRun.id == recon_id).first()
        if recon:
            recon.status = "error"
            recon.report_path = str(e)
            db.commit()
        from app.routers.error_logging import log_error
        log_error(db, None, "reconciliation", str(e), "")
    finally:
        db.close()


@router.post("/start", response_model=ReconciliationResponse)
async def start_reconciliation(
    data: ReconciliationRequest,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
):
    recon = ReconciliationRun(
        audit_type=data.audit_type,
        source_file_id=uuid.UUID(data.source_file_id),
        target_file_id=uuid.UUID(data.target_file_id) if data.target_file_id else None,
        output_path=data.output_path,
        user_id=uuid.UUID(data.user_id) if data.user_id else None,
    )
    db.add(recon)
    db.commit()
    db.refresh(recon)

    from app.routers.user_logs import log_activity
    log_activity(db, data.user_id, "reconciliation_start", "reconciliation", str(recon.id), f"Started: {data.audit_type}")

    from app.database import DATABASE_URL
    background_tasks.add_task(run_reconciliation_task, recon.id, DATABASE_URL)

    return recon


@router.get("/{recon_id}/status", response_model=ReconciliationResponse)
async def get_reconciliation_status(recon_id: uuid.UUID, db: Session = Depends(get_db)):
    recon = db.query(ReconciliationRun).filter(ReconciliationRun.id == recon_id).first()
    if not recon:
        raise HTTPException(status_code=404, detail="Reconciliation not found")
    return recon


@router.get("/{recon_id}/export")
async def export_results(recon_id: uuid.UUID, db: Session = Depends(get_db)):
    recon = db.query(ReconciliationRun).filter(ReconciliationRun.id == recon_id).first()
    if not recon:
        raise HTTPException(status_code=404, detail="Reconciliation not found")
    if not recon.report_path:
        raise HTTPException(status_code=404, detail="Output file not available")
    return FastAPIFileResponse(
        path=recon.report_path,
        filename=Path(recon.report_path).name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
