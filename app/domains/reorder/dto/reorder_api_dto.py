from pydantic import BaseModel, Field


class ReorderPlanResponse(BaseModel):
    summary: dict[str, int]
    rows: list[dict]


class ReorderPlanParams(BaseModel):
    w7: float = Field(default=0.5, ge=0.0, le=1.0)
    w14: float = Field(default=0.3, ge=0.0, le=1.0)
    w28: float = Field(default=0.2, ge=0.0, le=1.0)
    target_cover_days: int = Field(default=21, ge=1, le=120)
    use_dummy: bool = True

