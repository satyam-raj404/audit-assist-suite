from uuid import UUID
from datetime import datetime
from pydantic import BaseModel
from typing import Optional, List


# ─── Files ───

class FileResponse(BaseModel):
    id: UUID
    name: str
    size: int
    type: str

    class Config:
        from_attributes = True


# ─── PPT Automation (replaces Audit) ───

class PPTAutomationRequest(BaseModel):
    audit_type: str
    sub_audit_sector: str
    template_id: str
    output_path: str
    file_ids: List[str]
    user_id: Optional[str] = None


class AuditStepStatus(BaseModel):
    id: str
    label: str
    status: str


class PPTAutomationResponse(BaseModel):
    id: UUID
    status: str
    progress: float
    steps: List[AuditStepStatus] = []
    report_path: Optional[str] = None

    class Config:
        from_attributes = True


# ─── Reconciliation ───

class ReconciliationRequest(BaseModel):
    source_file_id: str
    target_file_id: str
    user_id: Optional[str] = None


class ReconciliationResponse(BaseModel):
    id: UUID
    status: str
    progress: float
    matched: int
    mismatched: int
    missing: int

    class Config:
        from_attributes = True


# ─── Auth ───

class RegisterRequest(BaseModel):
    name: str
    email: str
    password: str


class RegisterResponse(BaseModel):
    id: UUID
    name: str
    email: str

    class Config:
        from_attributes = True


class LoginRequest(BaseModel):
    email: str
    password: str


class LoginResponse(BaseModel):
    id: UUID
    name: str
    email: str

    class Config:
        from_attributes = True


# ─── Template Requests ───

class TemplateRequestCreate(BaseModel):
    spoc_email: str
    team_wide: bool = False
    user_id: Optional[str] = None


class TemplateRequestResponse(BaseModel):
    id: UUID
    file_name: str
    spoc_email: str
    team_wide: bool
    user_id: Optional[UUID] = None
    created_at: datetime

    class Config:
        from_attributes = True


# ─── User Profile (Onboarding) ───

class UserProfileCreate(BaseModel):
    user_id: str
    team_name: str
    partner_name: str
    client_name: str
    sector: str
    employee_id: str


class UserProfileResponse(BaseModel):
    id: UUID
    user_id: UUID
    team_name: str
    partner_name: str
    client_name: str
    sector: str
    employee_id: str
    created_at: datetime

    class Config:
        from_attributes = True


# ─── Error Logging ───

class ErrorLogCreate(BaseModel):
    user_id: Optional[str] = None
    module: str
    error_message: str
    stack_trace: Optional[str] = None


class ErrorLogResponse(BaseModel):
    id: UUID
    user_id: Optional[UUID] = None
    module: str
    error_message: str
    stack_trace: Optional[str] = None
    created_at: datetime

    class Config:
        from_attributes = True


# ─── User Logs ───

class UserLogCreate(BaseModel):
    user_id: Optional[str] = None
    action: str
    module: str
    run_id: Optional[str] = None
    details: Optional[str] = None


class UserLogResponse(BaseModel):
    id: UUID
    user_id: Optional[UUID] = None
    action: str
    module: str
    run_id: Optional[UUID] = None
    details: Optional[str] = None
    created_at: datetime

    class Config:
        from_attributes = True
