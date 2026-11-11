"""Cấu trúc dữ liệu chuẩn sau M0. Phân tích chỉ đọc các dataclass này."""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class ToKhai:
    loai: str  # gtgt | tndn | tncn | bctc
    ma_tkhai: str
    loai_tkhai: str  # C | B
    so_lan: int
    ky: str
    nam: int
    tu_ngay: str
    den_ngay: str
    mst: str
    ten_nnt: str
    chi_tieu: dict[str, float]
    file: str
    bctc_cdkt_cuoi: dict[str, float] = field(default_factory=dict)
    bctc_cdkt_dau: dict[str, float] = field(default_factory=dict)
    bctc_kqkd: dict[str, float] = field(default_factory=dict)


@dataclass
class HoaDon:
    chieu: str  # ban | mua
    ky: str
    so: str
    ngay: str
    mst_dt: str
    ten_dt: str
    mat_hang: str
    doanh_so: float
    thue_suat: float
    thue: float
    tong: float
    file: str
    sheet: str
    row: int


@dataclass
class Finding:
    id: str
    module: str  # m1 | m2 | m3
    tax_topic: str  # gtgt | tndn | tncn | hddt
    severity: str  # high | medium | low | skip
    title: str
    delta: str
    evidence: str
    legal_basis: str
    question_for_nnt: str
    tieu_chi: str = ""
    nam: int | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class HoSoChuan:
    mst: str
    ten_nnt: str
    nam: list[int]
    to_khai: list[ToKhai]
    hoa_don: list[HoaDon]
    skipped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
