from dataclasses import dataclass


@dataclass(frozen=True)
class ReorderSettings:
    w7: float
    w14: float
    w28: float
    target_cover_days: int

