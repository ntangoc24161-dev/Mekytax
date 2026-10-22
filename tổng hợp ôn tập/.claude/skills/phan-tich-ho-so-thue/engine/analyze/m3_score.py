"""M3: chấm điểm tổng hợp. Trọng số: high=3, medium=2, low=1 (chưa thấy bảng điểm VBA)."""
from __future__ import annotations

from models import Finding, HoSoChuan

W = {"high": 3, "medium": 2, "low": 1}


def run_m3(hs: HoSoChuan, findings: list[Finding]) -> list[Finding]:
    usable = [f for f in findings if f.severity in W]
    score = sum(W[f.severity] for f in usable)
    n_h = sum(1 for f in usable if f.severity == "high")
    n_m = sum(1 for f in usable if f.severity == "medium")
    n_l = sum(1 for f in usable if f.severity == "low")
    if n_h >= 3 or score >= 12:
        band, sev = "RỦI RO RẤT CAO", "high"
    elif n_h >= 1 or score >= 6:
        band, sev = "RỦI RO CAO", "high"
    elif score >= 3:
        band, sev = "RỦI RO", "medium"
    else:
        band, sev = "An toàn / ít dấu hiệu", "low"
    return [Finding(
        id="m3-score",
        module="m3",
        tax_topic="gtgt",
        severity=sev,
        title=f"Điểm rủi ro tổng thể: {band}",
        delta=f"Điểm {score} (high×3 + medium×2 + low×1). High={n_h}, medium={n_m}, low={n_l}. MST {hs.mst}.",
        evidence="Tổng hợp M1+M2",
        legal_basis="Thang điểm nội bộ (cần xác minh so với QĐ 1109/QĐ-CT 30/07/2026; chưa đọc được bảng điểm VBA)",
        question_for_nnt="Ưu tiên giải trình các mục High trước khi đoàn kiểm tra làm việc.",
        tieu_chi="SCORE",
    )]
