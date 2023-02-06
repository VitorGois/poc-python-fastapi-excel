from typing import List

from pydantic import BaseModel


class ExcelPositionDto(BaseModel):
    asset: str
    averagePrice: float
    quantity: float


class ExcelDto(BaseModel):
    positions: List[ExcelPositionDto] = []
