from pydantic import BaseModel, Field
from typing import Dict, Any

class DocumentGenerateRequest(BaseModel):
    template_id: str = Field(
        ...,
        description="Template identifier already uploaded",
        example="hos_AFS"
    )
    data: Dict[str, Any] = Field(
        ...,
        description="Key-value mapping for placeholders"
    )

class DocumentGenerateResponse(BaseModel):
    status: str = Field(example="success")
    replaced: int = Field(example=105)
    total: int = Field(example=113)
    output: str = Field(
        description="Generated DOCX file path",
        example="outputs/1_1770100369.docx"
    )
    generation_time_ms: int = Field(example=4135)
    trace_id: str = Field(example="d5b8f369")

