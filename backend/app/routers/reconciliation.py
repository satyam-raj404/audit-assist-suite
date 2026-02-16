import uuid
from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from app.database import get_db
from app.models import ReconciliationRun
from app.schemas import ReconciliationRequest, ReconciliationResponse
import time
import random
import io
import csv

router = APIRouter(prefix="/api/reconciliation", tags=["reconciliation"])


def run_reconciliation_task(recon_id: uuid.UUID, db_url: str):
    """Background task that simulates reconciliation.
    Replace with actual Excel comparison logic."""
    from sqlalchemy import create_engine
    from sqlalchemy.orm import sessionmaker

    engine = create_engine(db_url)
    Session = sessionmaker(bind=engine)
    db = Session()

    try:
        recon = db.query(ReconciliationRun).filter(ReconciliationRun.id == recon_id).first()
        if not recon:
            return

        recon.status = "running"
        db.commit()

        for i in range(10):
            time.sleep(0.5)
            recon.progress = ((i + 1) / 10) * 100
            db.commit()

        recon.matched = random.randint(200, 500)
        recon.mismatched = random.randint(10, 50)
        recon.missing = random.randint(5, 20)
        recon.status = "complete"
        recon.progress = 100
        db.commit()
    except Exception:
        recon.status = "error"
        db.commit()
    finally:
        db.close()


@router.post("/start", response_model=ReconciliationResponse)
async def start_reconciliation(
    data: ReconciliationRequest,
    background_tasks: BackgroundTasks,
    db: Session = Depends(get_db),
):
    recon = ReconciliationRun(
        source_file_id=uuid.UUID(data.source_file_id),
        target_file_id=uuid.UUID(data.target_file_id),
    )
    db.add(recon)
    db.commit()
    db.refresh(recon)

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

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Metric", "Count"])
    writer.writerow(["Matched", recon.matched])
    writer.writerow(["Mismatched", recon.mismatched])
    writer.writerow(["Missing", recon.missing])

    output.seek(0)
    return StreamingResponse(
        io.BytesIO(output.getvalue().encode()),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename=reconciliation_{recon_id}.csv"},
    )
