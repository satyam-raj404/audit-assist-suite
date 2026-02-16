from uuid import UUID
from datetime import datetime
from pydantic import BaseModel
from typing import Optional, List


# Files
class FileResponse(BaseModel):
    id: UUID
    name: str
    size: int
    type: str

    class Config:
        from_attributes = True


# Audit
class AuditRequest(BaseModel):
    audit_type: str
    sub_audit_sector: str
    template_id: str
    output_path: str
    file_ids: List[str]


class AuditStepStatus(BaseModel):
    id: str
    label: str
    status: str


class AuditResponse(BaseModel):
    id: UUID
    status: str
    progress: float
    steps: List[AuditStepStatus] = []
    report_path: Optional[str] = None

    class Config:
        from_attributes = True


# Reconciliation
class ReconciliationRequest(BaseModel):
    source_file_id: str
    target_file_id: str


class ReconciliationResponse(BaseModel):
    id: UUID
    status: str
    progress: float
    matched: int
    mismatched: int
    missing: int

    class Config:
        from_attributes = True
