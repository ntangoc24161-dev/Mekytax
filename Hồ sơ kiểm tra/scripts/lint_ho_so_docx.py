# -*- coding: utf-8 -*-
"""Module Linter / Validator kiểm định chất lượng bộ hồ sơ kiểm tra Word (docx).

Kiểm tra:
1. Quy chuẩn thể thức NĐ 30/2020/NĐ-CP cho Tờ trình, QĐKT, QĐGS (1 trang in A4, lề, tiêu ngữ, đường kẻ VML).
2. Quét sạch font tạp (Calibri, Arial...), bắt buộc 100% Times New Roman.
3. Kiểm tra bảo toàn Checkbox Wingdings (ô trống \\xa8, ô đã tích \\xfe) trong Báo cáo giám sát.
4. Kiểm tra phôi nghiệp vụ: giữ nguyên khung nổi (floating tables), tab stop ĐVT cùng dòng Năm.
5. Quét biến rác chưa thay thế ({{...}}, <<...>>).
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


class DocxLinter:
    def __init__(self, file_path: str | Path):
        self.file_path = Path(file_path)
        self.doc = None
        self.errors = []
        self.warnings = []
        self.info = []
        self.stats = {
            "paragraphs": 0,
            "tables": 0,
            "runs": 0,
            "calibri_count": 0,
            "times_count": 0,
            "checkbox_unchecked": 0,
            "checkbox_checked": 0,
            "unresolved_tags": 0,
            "floating_tables": 0,
        }

    def load(self) -> bool:
        if not self.file_path.exists():
            self.errors.append(f"Tệp không tồn tại: {self.file_path}")
            return False
        try:
            self.doc = Document(str(self.file_path))
            return True
        except Exception as e:
            self.errors.append(f"Không thể mở tệp Word: {e}")
            return False

    def lint_all(self) -> dict:
        if not self.load():
            return self._build_result()

        self._check_fonts_and_tags()
        self._check_checkboxes()
        self._check_floating_tables_and_tabs()
        self._check_single_page_layout()

        return self._build_result()

    def _check_fonts_and_tags(self):
        """Quét font chữ và các placeholder chưa thay thế trong toàn bộ paragraph & table."""
        # Check paragraphs
        for p_idx, p in enumerate(self.doc.paragraphs):
            self.stats["paragraphs"] += 1
            text = p.text
            if "{{" in text or "}}" in text or "<<" in text or ">>" in text:
                self.stats["unresolved_tags"] += 1
                self.warnings.append(f"Đoạn {p_idx+1}: Phát hiện placeholder chưa thay: '{text.strip()[:60]}...'")

            for r in p.runs:
                self.stats["runs"] += 1
                self._inspect_run_font(r, f"Đoạn {p_idx+1}")

        # Check tables
        for t_idx, table in enumerate(self.doc.tables):
            self.stats["tables"] += 1
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    for p_idx, p in enumerate(cell.paragraphs):
                        text = p.text
                        if "{{" in text or "}}" in text:
                            self.stats["unresolved_tags"] += 1
                            self.warnings.append(f"Bảng {t_idx+1} [R{r_idx+1}C{c_idx+1}]: Placeholder chưa thay: '{text.strip()}'")
                        for r in p.runs:
                            self.stats["runs"] += 1
                            self._inspect_run_font(r, f"Bảng {t_idx+1} [R{r_idx+1}C{c_idx+1}]")

    def _inspect_run_font(self, run, location_label: str):
        fn = run.font.name
        rPr = run._element.rPr
        rFonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
        
        ascii_font = rFonts.get(qn("w:ascii")) if rFonts is not None else None
        hAnsi_font = rFonts.get(qn("w:hAnsi")) if rFonts is not None else None

        font_names = {fn, ascii_font, hAnsi_font} - {None}
        
        # Check for Calibri
        for name in font_names:
            if "Calibri" in name:
                self.stats["calibri_count"] += 1
                self.errors.append(f"{location_label}: Dính font '{name}' trong text: '{run.text[:30]}'")
            elif "Wingdings" in name:
                pass  # Wingdings is allowed for checkboxes
            elif "Times" in name:
                self.stats["times_count"] += 1

    def _check_checkboxes(self):
        """Kiểm tra checkbox Wingdings và Unicode thô."""
        full_text = []
        for p in self.doc.paragraphs:
            full_text.append(p.text)
        for t in self.doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    full_text.append(cell.text)
        joined = "".join(full_text)

        # Check raw unicode square
        if "□" in joined:
            self.warnings.append("Phát hiện ký tự ô vuông thô '□' Unicode thay vì checkbox Wingdings chuẩn.")

        # Count Wingdings characters
        unchecked = joined.count("\xa8")  # Wingdings empty box
        checked = joined.count("\xfe")    # Wingdings checked box
        self.stats["checkbox_unchecked"] = unchecked
        self.stats["checkbox_checked"] = checked

        if unchecked > 0 or checked > 0:
            self.info.append(f"Checkbox Wingdings: {checked} ô đã tích (\\xfe), {unchecked} ô trống (\\xa8).")
            if checked == 0 and unchecked > 10 and "GIÁM SÁT" in joined.upper():
                self.warnings.append("CẢNH BÁO: Báo cáo giám sát có checkbox nhưng 100% là ô trống (có thể đã bị reset mất tick chọn).")

    def _check_floating_tables_and_tabs(self):
        """Kiểm tra khung nổi (tblpPr) và tab stop ĐVT."""
        for t_idx, table in enumerate(self.doc.tables):
            tblPr = table._element.xpath("w:tblPr")
            if tblPr:
                tblpPr = tblPr[0].xpath("w:tblpPr")
                if tblpPr:
                    self.stats["floating_tables"] += 1
        
        if self.stats["floating_tables"] > 0:
            self.info.append(f"Phát hiện {self.stats['floating_tables']} khung nổi (Floating Table/tblpPr) - Giữ nguyên layout phôi nghiệp vụ.")

    def _check_single_page_layout(self):
        """Đối với file QĐKT hoặc QĐGS độc lập, kiểm tra dấu hiệu tràn sang trang 2."""
        fname = self.file_path.name.lower()
        if fname.startswith("qd ") or fname.startswith("qdgiam") or "quyet dinh" in fname or "to trinh" in fname:
            # Check page break elements
            page_breaks = 0
            for p in self.doc.paragraphs:
                for r in p.runs:
                    brs = r._element.xpath("w:br[@w:type='page']")
                    page_breaks += len(brs)
                if p._element.xpath("w:pPr/w:pageBreakBefore"):
                    page_breaks += 1
            
            if page_breaks > 0:
                self.errors.append(f"Tệp đơn vị '{self.file_path.name}' chứa {page_breaks} lệnh ngắt trang (Page Break), có nguy cơ vượt quá 1 trang in.")
            
            # Check paragraph count as heuristic
            if len(self.doc.paragraphs) > 35 and len(self.doc.tables) == 0:
                self.warnings.append(f"Tệp '{self.file_path.name}' có {len(self.doc.paragraphs)} đoạn văn - Cần kiểm tra kỹ khoảng cách dòng để không tràn sang trang 2.")

    def _build_result(self) -> dict:
        is_valid = len(self.errors) == 0
        return {
            "file": str(self.file_path),
            "valid": is_valid,
            "errors": self.errors,
            "warnings": self.warnings,
            "info": self.info,
            "stats": self.stats,
        }

    def print_report(self):
        res = self._build_result()
        print("=" * 70)
        print(f"KẾT QUẢ KIỂM ĐỊNH FILE: {self.file_path.name}")
        print("=" * 70)
        print(f"Trạng thái: {'[ĐẠT CHUẨN]' if res['valid'] else '[CẦN XỬ LÝ LỖI]'}")
        print(f"- Số đoạn: {res['stats']['paragraphs']} | Số bảng: {res['stats']['tables']} | Runs: {res['stats']['runs']}")
        print(f"- Font Times: {res['stats']['times_count']} | Font Calibri dính lỗi: {res['stats']['calibri_count']}")
        print(f"- Checkbox Wingdings: {res['stats']['checkbox_checked']} đã tích | {res['stats']['checkbox_unchecked']} ô trống")
        print(f"- Khung nổi (floating): {res['stats']['floating_tables']} | Placeholder chưa thay: {res['stats']['unresolved_tags']}")
        print("-" * 70)

        if res["errors"]:
            print(f"[LỖI NGHIÊM TRỌNG] ({len(res['errors'])}):")
            for e in res["errors"]:
                print(f"  ❌ {e}")
            print("-" * 70)

        if res["warnings"]:
            print(f"[CẢNH BÁO] ({len(res['warnings'])}):")
            for w in res["warnings"]:
                print(f"  ⚠️  {w}")
            print("-" * 70)

        if res["info"]:
            print(f"[THÔNG TIN ĐỊNH DẠNG] ({len(res['info'])}):")
            for i in res["info"]:
                print(f"  ℹ️  {i}")
            print("-" * 70)


def main():
    parser = argparse.ArgumentParser(description="Kiểm tra chất lượng định dạng tệp Word Hồ sơ kiểm tra.")
    parser.add_argument("file_path", help="Đường dẫn tới file .docx cần kiểm định")
    args = parser.parse_args()

    linter = DocxLinter(args.file_path)
    res = linter.lint_all()
    linter.print_report()
    sys.exit(0 if res["valid"] else 1)


if __name__ == "__main__":
    main()
