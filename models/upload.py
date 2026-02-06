from pydantic import BaseModel, Field 
from typing import Dict, Any

class DocumentUploadRequest(BaseModel):
    template_id: str = Field(..., example="hos_AFS.docx")
    data: Dict[str:Any]

class DocumentUploadResponse(BaseModel):
    status: str
    replaced: int
    total: int
    file_url: str
    generation_time_ms: int
    trace_id: str

