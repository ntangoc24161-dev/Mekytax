"""JSON findings -> Markdown ngữ cảnh cao cho Javis."""
from __future__ import annotations

from datetime import date
from pathlib import Path

from models import Finding, HoSoChuan

SEV_VI = {"high": "Cao", "medium": "Trung bình", "low": "Thấp"}
TOPIC_H = {
    "gtgt": "1. Thuế GTGT",
    "tndn": "2. Thuế TNDN",
    "tncn": "3. Thuế TNCN",
    "hddt": "4. Hóa đơn rủi ro",
}


def wrap_source(body: str, hs: HoSoChuan, n: int) -> str:
    today = date.today().isoformat()
    nam = "-".join(str(y) for y in hs.nam) if hs.nam else ""
    fm = "\n".join([
        "---",
        "type: source",
        "source_kind: tax-analysis",
        "status: unprocessed",
        f"created: {today}",
        f'title: "Phân tích hồ sơ thuế {hs.mst} {nam}"',
        f"mst: \"{hs.mst}\"",
        f"ky: \"{nam}\"",
        f"findings: {n}",
        "tags: [thue, phan-tich, rui-ro]",
        "---",
        "",
    ])
    return fm + body


def to_markdown(hs: HoSoChuan, findings: list[Finding]) -> str:
    usable = [f for f in findings if f.severity != "skip"]
    n_h = sum(1 for f in usable if f.severity == "high")
    n_m = sum(1 for f in usable if f.severity == "medium")
    n_l = sum(1 for f in usable if f.severity == "low")
    score = next((f for f in usable if f.id == "m3-score"), None)
    lines = [
        f"# Phân tích hồ sơ kiểm tra thuế - {hs.ten_nnt or hs.mst}",
        "",
        f"MST **{hs.mst}**. Kỳ **{', '.join(str(y) for y in hs.nam) or '(không rõ)'}**. "
        f"{len(hs.to_khai)} tờ khai/BCTC XML, {len(hs.hoa_don)} dòng hóa đơn.",
        "",
        "## Tóm tắt rủi ro",
        "",
        f"Phát hiện **{len(usable)}** mục: **{n_h} cao**, **{n_m} trung bình**, **{n_l} thấp**.",
    ]
    if score:
        lines.append("")
        lines.append(score.delta)
    highs = [f for f in usable if f.severity == "high" and f.id != "m3-score"][:5]
    if highs:
        lines += ["", "Ưu tiên giải trình:"]
        for f in highs:
            lines.append(f"- {f.title} ({f.delta})")
    lines += ["", "Ngưỡng lấy từ sheet TUDIEN tool 4.4. Tiêu chí cần TMS (TC24) không chạy."]
    for topic, heading in TOPIC_H.items():
        group = [f for f in usable if f.tax_topic == topic and f.id != "m3-score"]
        lines += ["", f"# {heading}", ""]
        if not group:
            lines.append("_Không phát hiện mục thuộc chuyên đề này._")
            continue
        for f in group:
            lines += _block(f)
    if hs.warnings or hs.skipped:
        lines += ["", "# Phụ lục ingest", ""]
        for w in hs.warnings[:20]:
            lines.append(f"- Cảnh báo: {w}")
        for s in hs.skipped[:30]:
            lines.append(f"- Bỏ: {s}")
    text = "\n".join(lines).strip() + "\n"
    if "\u2014" in text:
        text = text.replace("\u2014", "-")
    return wrap_source(text, hs, len(usable))


def _block(f: Finding) -> list[str]:
    sev = SEV_VI.get(f.severity, f.severity)
    tc = f" `{f.tieu_chi}`" if f.tieu_chi else ""
    return [
        f"### {f.title}{tc} ({sev})",
        "",
        f"- **Nội dung bất thường:** {f.delta}",
        f"- **Căn cứ dữ liệu:** {f.evidence}",
        f"- **Căn cứ tiêu chí:** {f.legal_basis}",
        f"- **Gợi ý câu hỏi kiểm tra:** {f.question_for_nnt}",
        "",
    ]


def write_outputs(out_dir: Path, hs: HoSoChuan, findings: list[Finding], md: str) -> tuple[Path, Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = f"phan-tich-{hs.mst or 'nnt'}-{'-'.join(str(y) for y in hs.nam) or 'ky'}"
    md_path = out_dir / f"{stem}.md"
    js_path = out_dir / f"{stem}.json"
    xlsx_path = out_dir / f"{stem}.xlsx"
    import json
    from report.excel import write_xlsx
    js_path.write_text(
        json.dumps(
            {"mst": hs.mst, "ten": hs.ten_nnt, "nam": hs.nam,
             "findings": [f.to_dict() for f in findings]},
            ensure_ascii=False, indent=2,
        ),
        encoding="utf-8",
    )
    md_path.write_text(md, encoding="utf-8")
    write_xlsx(xlsx_path, hs, findings)
    return md_path, js_path, xlsx_path
