import os
from pathlib import Path
from fastapi import APIRouter, Query
from typing import List, Optional
from pydantic import BaseModel

router = APIRouter(prefix="/api/ppt-templates", tags=["ppt_templates"])

# Root folder where all template .pptx files are stored
TEMPLATES_ROOT = Path(__file__).resolve().parents[2] / "templates"


class PptTemplate(BaseModel):
    id: str
    name: str
    description: str
    audit_type: str
    utility_type: str
    path: str
    available: bool  # True only if the .pptx file actually exists on disk


# Registry of all known templates — one entry per audit_type + utility_type combination
_REGISTRY: List[dict] = [
    {
        "id": "ia_standard",
        "name": "IA Standard Template",
        "description": (
            "Standard Internal Audit report template. Contains observation slides, "
            "executive summary, and risk rating layout. Compatible with the IA issue tracker."
        ),
        "audit_type": "Internal Audit",
        "utility_type": "Report",
        "file": "IA_Standard/template.pptx",
    },
    {
        "id": "ia_zensar",
        "name": "IA Zensar Template",
        "description": (
            "Internal Audit report template for Zensar engagements. Uses placeholder-based "
            "layout with customised branding and observation structure."
        ),
        "audit_type": "Internal Audit-Zensar",
        "utility_type": "Report",
        "file": "IA_Zensar/template.pptx",
    },
    {
        "id": "concurrent_report",
        "name": "Concurrent Audit Report Template",
        "description": (
            "Concurrent Audit report template supporting Draft and Final versions. "
            "Includes branch-level summary, transaction observations, and volume data slides."
        ),
        "audit_type": "Concurrent Audit",
        "utility_type": "Report",
        "file": "Concurrent_Report/template.pptx",
    },
    {
        "id": "concurrent_dashboard",
        "name": "Concurrent Audit Dashboard Template",
        "description": (
            "Dashboard template for Concurrent Audit. Contains KPI charts, trend graphs, "
            "and compliance summary slides populated from the issue tracker."
        ),
        "audit_type": "Concurrent Audit",
        "utility_type": "Dashboard",
        "file": "Concurrent_Dashboard/template.pptx",
    },
    {
        "id": "icofr_dashboard",
        "name": "ICOFR Dashboard Template",
        "description": (
            "ICOFR (Internal Controls over Financial Reporting) dashboard template. "
            "Covers control testing results, deficiency classification, and remediation tracking."
        ),
        "audit_type": "ICOFR",
        "utility_type": "Dashboard",
        "file": "ICOFR_Dashboard/template.pptx",
    },
]


def _build(entry: dict) -> PptTemplate:
    abs_path = str(TEMPLATES_ROOT / entry["file"])
    return PptTemplate(
        id=entry["id"],
        name=entry["name"],
        description=entry["description"],
        audit_type=entry["audit_type"],
        utility_type=entry["utility_type"],
        path=abs_path,
        available=os.path.isfile(abs_path),
    )


@router.get("", response_model=List[PptTemplate])
async def list_templates(
    audit_type: Optional[str] = Query(None),
    utility_type: Optional[str] = Query(None),
):
    results = _REGISTRY
    if audit_type:
        results = [t for t in results if t["audit_type"] == audit_type]
    if utility_type:
        results = [t for t in results if t["utility_type"] == utility_type]
    return [_build(t) for t in results]


@router.get("/{template_id}", response_model=PptTemplate)
async def get_template(template_id: str):
    for entry in _REGISTRY:
        if entry["id"] == template_id:
            return _build(entry)
    from fastapi import HTTPException
    raise HTTPException(status_code=404, detail="Template not found")
