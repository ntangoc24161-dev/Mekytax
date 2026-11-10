# -*- coding: utf-8 -*-
"""OCR tieng Viet offline: PP-OCRv5 det + latin_PP-OCRv5_mobile_rec (giu dau)."""
import sys, os, io, time, argparse, tempfile

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("out")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--first", type=int, default=1)
    ap.add_argument("--last", type=int, default=0)
    a = ap.parse_args()

    import fitz
    from paddleocr import PaddleOCR

    doc = fitz.open(a.pdf)
    n = doc.page_count
    last = a.last if a.last > 0 else n
    print(f"[i] {os.path.basename(a.pdf)}: {n} trang, OCR trang {a.first}-{last}, dpi={a.dpi}", flush=True)

    ocr = PaddleOCR(
        text_detection_model_name="PP-OCRv5_server_det",
        text_recognition_model_name="latin_PP-OCRv5_mobile_rec",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        lang="latin",
    )
    print("[i] model san sang", flush=True)

    tmpdir = tempfile.mkdtemp(prefix="ocrvi_")
    chunks = []
    for i in range(a.first - 1, last):
        t0 = time.time()
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=a.dpi)
        img = os.path.join(tmpdir, f"p{i+1:03d}.png")
        pix.save(img)
        res = ocr.predict(img)
        lines = []
        for r in res:
            d = r.json.get("res", r.json) if hasattr(r, "json") else r
            texts = d.get("rec_texts") or []
            boxes = d.get("rec_polys") or d.get("dt_polys") or []
            items = []
            for k, t in enumerate(texts):
                if not t or not t.strip():
                    continue
                y = 0
                x = 0
                if k < len(boxes):
                    b = boxes[k]
                    try:
                        ys = [float(p[1]) for p in b]
                        xs = [float(p[0]) for p in b]
                        y = sum(ys) / len(ys)
                        x = min(xs)
                    except Exception:
                        pass
                items.append((y, x, t.strip()))
            items.sort(key=lambda z: (round(z[0] / 12), z[1]))
            cur_y = None
            buf = []
            for y, x, t in items:
                if cur_y is None or abs(y - cur_y) <= 12:
                    buf.append(t)
                    cur_y = y if cur_y is None else cur_y
                else:
                    lines.append(" ".join(buf))
                    buf = [t]
                    cur_y = y
            if buf:
                lines.append(" ".join(buf))
        txt = "\n".join(lines)
        chunks.append(f"\n\n---\n\n## Trang {i+1}\n\n{txt}")
        print(f"[i] trang {i+1}/{last}: {len(txt)} ky tu, {time.time()-t0:.1f}s", flush=True)
        try:
            os.remove(img)
        except Exception:
            pass

    body = "".join(chunks)
    head = (
        "---\n"
        "type: source\n"
        "source_kind: screenshot\n"
        "status: unprocessed\n"
        "created: 2026-11-09\n"
        f"ocr_engine: PaddleOCR 3.7.0 det=PP-OCRv5_server_det rec=latin_PP-OCRv5_mobile_rec dpi={a.dpi}\n"
        f"source_file: {a.pdf}\n"
        f"pages: {last - a.first + 1}\n"
        "---\n\n"
        f"# OCR {os.path.basename(a.pdf)}\n"
    )
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with io.open(a.out, "w", encoding="utf-8") as f:
        f.write(head + body + "\n")
    print(f"[OK] ghi {a.out} ({len(head)+len(body)} ky tu)", flush=True)

main()
