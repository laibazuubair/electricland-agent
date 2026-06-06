from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime
import uuid

# ── Individual extracted field ────────────────────────────────────────────────
class ExtractedField(BaseModel):
    value: Optional[float | str] = None
    confidence: str = "low"  # high / medium / low

class UnitField(BaseModel):
    unit_name: str
    annual_rent_gbp: Optional[float] = None

class UnitsField(BaseModel):
    value: Optional[List[UnitField]] = None
    confidence: str = "low"

# ── Full extracted deal data ──────────────────────────────────────────────────
class ExtractedDealData(BaseModel):
    site_address: ExtractedField
    site_area_acres: ExtractedField
    purchase_price_gbp: ExtractedField
    build_cost_gbp: ExtractedField
    total_development_cost_gbp: ExtractedField
    total_rent_roll_pa_gbp: ExtractedField
    individual_units: UnitsField
    target_yield_percent: ExtractedField
    gdv_gbp: ExtractedField
    profit_on_cost_percent: ExtractedField
    development_timeline_months: ExtractedField
    planning_status: ExtractedField
    sender_name: ExtractedField
    flags: List[str] = []

# ── Deal record stored in database ───────────────────────────────────────────
class Deal(BaseModel):
    id: str = None
    created_at: str = None
    status: str = "pending_review"  # pending_review / approved / amended
    email_subject: str = ""
    email_from: str = ""
    email_body: str = ""
    extracted_data: Optional[ExtractedDealData] = None
    excel_filename: str = ""
    approved_by: Optional[str] = None
    approved_at: Optional[str] = None
    analyst_notes: Optional[str] = None

    def __init__(self, **data):
        super().__init__(**data)
        if not self.id:
            self.id = str(uuid.uuid4())[:8]
        if not self.created_at:
            self.created_at = datetime.now().strftime("%d %b %Y %H:%M:%S")

# ── Request bodies for API endpoints ─────────────────────────────────────────
class ProcessEmailRequest(BaseModel):
    subject: str
    email_from: str
    body: str

class ApproveRequest(BaseModel):
    approved_by: str
    notes: Optional[str] = ""

class AmendRequest(BaseModel):
    amended_by: str
    notes: Optional[str] = ""
    amendments: dict  # field name → new value