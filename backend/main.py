from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, FileResponse
from fastapi.background import BackgroundTasks
import fitz  # PyMuPDF
import io
import json
import os
import re
import tempfile
import base64
import traceback
import subprocess
import numpy as np
from PIL import Image

import docx
from docx import Document as DocxDocument
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor

import xlsxwriter

# Aumentar límite de campo form para python-multipart (permite imágenes grandes en base64)
try:
    import multipart.multipart as _mp
    _mp.MAX_FORM_MEMORY_SIZE = 100 * 1024 * 1024   # 100 MB
except Exception:
    pass

try:
    from starlette.formparsers import MultiPartParser
    MultiPartParser.max_file_size = 100 * 1024 * 1024  # 100 MB
except Exception:
    pass

app = FastAPI(title="PDF & Image Admin Converter API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp", "tiff", "tif"}


@app.get("/")
def read_root():
    return {"status": "ok"}


# =====================================================================
# 1. EDITOR DE PDF ENDPOINT
# =====================================================================
@app.post("/api/pdf/edit")
async def edit_pdf(request: Request):
    try:
        form = await request.form(max_files=100, max_fields=100)
        file_upload = form.get("file")
        actions_str = form.get("actions", "[]")
        if file_upload is None:
            raise HTTPException(status_code=400, detail="No se recibió el archivo PDF.")
        pdf_bytes = await file_upload.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        action_list = json.loads(actions_str)

        for action in action_list:
            page_num = action.get("page", 1) - 1
            if page_num < 0 or page_num >= len(doc):
                continue

            page = doc[page_num]
            action_type = action.get("type")

            if action_type == "text":
                text = action.get("text", "")
                x = action.get("x", 0)
                y = action.get("y", 0)
                size = action.get("size", 12)
                color_hex = action.get("color", "#000000")
                is_bold = action.get("bold", False)

                try:
                    color_hex = color_hex.lstrip("#")
                    r = int(color_hex[0:2], 16) / 255.0
                    g = int(color_hex[2:4], 16) / 255.0
                    b = int(color_hex[4:6], 16) / 255.0
                    color_tuple = (r, g, b)
                except Exception:
                    color_tuple = (0, 0, 0)

                base_font = action.get("font", "helv")
                if is_bold:
                    if base_font == "helv":
                        fontname = "hebo"
                    elif base_font == "ti-ro":
                        fontname = "tibo"
                    elif base_font == "cour":
                        fontname = "cobo"
                    else:
                        fontname = "hebo"
                else:
                    fontname = base_font

                point = fitz.Point(x, y)
                page.insert_text(point, text, fontsize=size, color=color_tuple, fontname=fontname)

            elif action_type == "image":
                src = action.get("src", "")
                if src.startswith("data:image"):
                    try:
                        base64_data = src.split(",")[1]
                        image_bytes = base64.b64decode(base64_data)

                        x = action.get("x", 0)
                        y = action.get("y", 0)
                        w = action.get("width", 100)
                        h = action.get("height", 100)

                        rect = fitz.Rect(x, y, x + w, y + h)
                        page.insert_image(rect, stream=image_bytes)
                    except Exception as e:
                        print(f"Error inserting image: {e}")

            elif action_type == "watermark":
                text = action.get("text", "CONFIDENTIAL")
                rect = page.rect
                page.insert_text(
                    fitz.Point(rect.width / 4, rect.height / 2),
                    text,
                    fontsize=50,
                    color=(0.5, 0.5, 0.5),
                    rotate=45,
                    fill_opacity=0.3,
                )

        pdf_bytes_out = doc.tobytes(garbage=4, deflate=True)
        doc.close()
        return Response(
            content=pdf_bytes_out,
            media_type="application/pdf",
            headers={"Content-Disposition": 'attachment; filename="PdfEditado.pdf"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        tb = traceback.format_exc()
        print(f"\n=== ERROR EN /api/pdf/edit ===\n{tb}\n=============================")
        raise HTTPException(status_code=500, detail=str(e))


# =====================================================================
# 2. WINDOWS NATIVE OCR & IMAGE TABLE DETECTION
# =====================================================================
def run_windows_ocr(image_path: str):
    """
    Runs Windows built-in OcrEngine (Windows.Media.Ocr) on an image file.
    Returns a list of word tuples: (x0, y0, x1, y1, text) in the original image's coordinate space.
    """
    try:
        with Image.open(image_path) as im:
            orig_w, orig_h = im.size
            scale = 1.0
            # Upscale smaller images so OCR accurately detects small table text
            if orig_w < 1400:
                scale = 2.0
                new_size = (int(orig_w * scale), int(orig_h * scale))
                upscaled = im.convert("RGB").resize(new_size, Image.Resampling.LANCZOS)
                fd_tmp, ocr_img_path = tempfile.mkstemp(suffix=".png")
                os.close(fd_tmp)
                upscaled.save(ocr_img_path, "PNG")
            elif im.mode not in ("RGB", "RGBA", "L"):
                fd_tmp, ocr_img_path = tempfile.mkstemp(suffix=".png")
                os.close(fd_tmp)
                im.convert("RGB").save(ocr_img_path, "PNG")
            else:
                ocr_img_path = image_path
    except Exception as e:
        print(f"Error preparing image for OCR: {e}")
        return []

    ps_script_content = r"""
param([string]$ImagePath)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]

    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]

    function Await($WinRtTask, $ResultType) {
        $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
        $netTask = $asTask.Invoke($null, @($WinRtTask))
        $netTask.Wait(-1) | Out-Null
        return $netTask.Result
    }

    $fs = [System.IO.File]::OpenRead($ImagePath)
    $stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($fs)
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        $langs = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
        if ($langs.Count -gt 0) {
            $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($langs[0])
        }
    }
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    $words = @()
    foreach ($line in $result.Lines) {
        foreach ($w in $line.Words) {
            $words += @{
                t = $w.Text
                x = [double]$w.BoundingRect.X
                y = [double]$w.BoundingRect.Y
                w = [double]$w.BoundingRect.Width
                h = [double]$w.BoundingRect.Height
            }
        }
    }
    $stream.Dispose()
    $fs.Dispose()
    ConvertTo-Json -InputObject @($words) -Depth 4 -Compress
} catch {
    Write-Output "[]"
}
"""
    fd_ps, ps_path = tempfile.mkstemp(suffix=".ps1")
    with os.fdopen(fd_ps, "w", encoding="utf-8") as f:
        f.write(ps_script_content)

    words_out = []
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path, "-ImagePath", os.path.abspath(ocr_img_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )
        stdout = (res.stdout or "").strip()
        if stdout:
            idx = stdout.find("[")
            if idx != -1:
                stdout = stdout[idx:]
            parsed = json.loads(stdout)
            if isinstance(parsed, dict):
                parsed = [parsed]
            for item in parsed:
                t = str(item.get("t", "")).strip()
                if not t:
                    continue
                x0 = float(item.get("x", 0)) / scale
                y0 = float(item.get("y", 0)) / scale
                w = float(item.get("w", 0)) / scale
                h = float(item.get("h", 0)) / scale
                words_out.append((x0, y0, x0 + w, y0 + h, t))
    except Exception as e:
        print(f"Error in run_windows_ocr: {e}")
    finally:
        cleanup_files([ps_path])
        if ocr_img_path != image_path:
            cleanup_files([ocr_img_path])

    return words_out


def clean_table_rows(rows):
    """Cleans a 2D table array, normalizing strings and stripping completely empty rows/columns."""
    if not rows:
        return []
    cleaned = []
    max_cols = max((len(r) for r in rows), default=0)
    if max_cols == 0:
        return []

    for r in rows:
        norm_row = []
        for c_idx in range(max_cols):
            val = r[c_idx] if c_idx < len(r) else ""
            if val is None:
                val = ""
            val = re.sub(r"\s+", " ", str(val)).strip()
            norm_row.append(val)
        if any(cell != "" for cell in norm_row):
            cleaned.append(norm_row)

    if not cleaned:
        return []

    # Remove completely empty columns
    non_empty_cols = [
        c_idx for c_idx in range(max_cols)
        if any(row[c_idx] != "" for row in cleaned)
    ]
    if not non_empty_cols:
        return []

    final_rows = [
        [row[c_idx] for c_idx in non_empty_cols]
        for row in cleaned
    ]
    return final_rows


def cluster_1d_positions(positions, tolerance=6.0):
    """Clusters sorted 1D coordinates that are within `tolerance` of each other."""
    if not positions:
        return []
    positions = sorted(positions)
    clusters = [[positions[0]]]
    for pos in positions[1:]:
        if pos - clusters[-1][-1] <= tolerance:
            clusters[-1].append(pos)
        else:
            clusters.append([pos])
    return [sum(c) / len(c) for c in clusters]


def detect_grid_tables_in_image(image_path: str, ocr_words):
    """
    Detects tables with visible horizontal/vertical grid lines in an image using numpy,
    and populates their cells using OCR word coordinates.
    """
    try:
        with Image.open(image_path) as im:
            gray = np.array(im.convert("L"), dtype=np.float32)
    except Exception:
        return []

    h, w = gray.shape
    if h < 40 or w < 40:
        return []

    # Detect lines (both dark lines on light background and light lines on dark background)
    median_bg = np.median(gray)
    if median_bg > 127:
        binary = gray < (median_bg - 30)
    else:
        binary = gray > (median_bg + 30)

    # Find horizontal line candidates: rows where a significant segment of pixels is active
    min_h_len = max(int(w * 0.15), 50)
    min_v_len = max(int(h * 0.08), 35)

    # Check horizontal continuity using sliding window / run length approximation
    h_line_ys = []
    h_spans = []
    for y in range(h):
        row = binary[y]
        if np.sum(row) < min_h_len:
            continue
        # Find longest contiguous run (allowing up to 2px gap)
        padded = np.convolve(row.astype(int), np.ones(3, dtype=int), mode="same") > 0
        diffs = np.diff(np.concatenate(([0], padded.astype(int), [0])))
        starts = np.where(diffs == 1)[0]
        ends = np.where(diffs == -1)[0]
        for s, e in zip(starts, ends):
            if (e - s) >= min_h_len:
                h_line_ys.append(y)
                h_spans.append((s, e))
                break

    v_line_xs = []
    v_spans = []
    for x in range(w):
        col = binary[:, x]
        if np.sum(col) < min_v_len:
            continue
        padded = np.convolve(col.astype(int), np.ones(3, dtype=int), mode="same") > 0
        diffs = np.diff(np.concatenate(([0], padded.astype(int), [0])))
        starts = np.where(diffs == 1)[0]
        ends = np.where(diffs == -1)[0]
        for s, e in zip(starts, ends):
            if (e - s) >= min_v_len:
                v_line_xs.append(x)
                v_spans.append((s, e))
                break

    y_grid = cluster_1d_positions(h_line_ys, tolerance=max(4.0, h * 0.008))
    x_grid = cluster_1d_positions(v_line_xs, tolerance=max(4.0, w * 0.008))

    if len(y_grid) < 3 or len(x_grid) < 3:
        return []

    # Split y_grid into separate vertical table groups if vertical lines do not bridge adjacent y_grid levels
    y_groups = [[y_grid[0]]]
    for idx in range(len(y_grid) - 1):
        y_top = y_grid[idx]
        y_bot = y_grid[idx + 1]
        mid_y = (y_top + y_bot) / 2.0
        # Check if at least 2 vertical line spans cover mid_y
        bridging_v = sum(1 for (vs, ve) in v_spans if (vs - 5) <= mid_y <= (ve + 5))
        if bridging_v >= 2:
            y_groups[-1].append(y_bot)
        else:
            y_groups.append([y_bot])

    detected_tables = []
    for yg in y_groups:
        if len(yg) < 3:
            continue
        num_rows = len(yg) - 1
        num_cols = len(x_grid) - 1
        if num_rows < 2 or num_cols < 2:
            continue

        grid_bbox = (x_grid[0], yg[0], x_grid[-1], yg[-1])
        cells = [[[] for _ in range(num_cols)] for _ in range(num_rows)]

        for wx0, wy0, wx1, wy1, text in ocr_words:
            cx = (wx0 + wx1) / 2.0
            cy = (wy0 + wy1) / 2.0
            if not (grid_bbox[0] - 5 <= cx <= grid_bbox[2] + 5 and grid_bbox[1] - 5 <= cy <= grid_bbox[3] + 5):
                continue
            r_idx = None
            for r in range(num_rows):
                if yg[r] - 3 <= cy <= yg[r + 1] + 3:
                    r_idx = r
                    break
            c_idx = None
            for c in range(num_cols):
                if x_grid[c] - 3 <= cx <= x_grid[c + 1] + 3:
                    c_idx = c
                    break
            if r_idx is not None and c_idx is not None:
                cells[r_idx][c_idx].append((wy0, wx0, text))

        table_rows = []
        for r in range(num_rows):
            row_strings = []
            for c in range(num_cols):
                cell_words = sorted(cells[r][c], key=lambda item: (round(item[0] / 6.0), item[1]))
                row_strings.append(" ".join(w_item[2] for w_item in cell_words))
            table_rows.append(row_strings)

        cleaned = clean_table_rows(table_rows)
        if len(cleaned) >= 2 and len(cleaned[0]) >= 2:
            detected_tables.append({"bbox": grid_bbox, "rows": cleaned})

    return detected_tables


def extract_tables_from_words(words, page_width: float, page_height: float, exclude_bboxes=None):
    """
    Detects tables (including borderless tables) from word bounding boxes:
    words: list of (x0, y0, x1, y1, text)
    Returns list of {"bbox": (x0, y0, x1, y1), "rows": [[...], ...]}.
    """
    if not words:
        return []

    exclude_bboxes = exclude_bboxes or []

    def is_excluded(wx0, wy0, wx1, wy1):
        cx = (wx0 + wx1) / 2.0
        cy = (wy0 + wy1) / 2.0
        for bx0, by0, bx1, by1 in exclude_bboxes:
            if bx0 - 2 <= cx <= bx1 + 2 and by0 - 2 <= cy <= by1 + 2:
                return True
        return False

    filtered = [w for w in words if not is_excluded(w[0], w[1], w[2], w[3]) and str(w[4]).strip()]
    if len(filtered) < 4:
        return []

    heights = [max(2.0, w[3] - w[1]) for w in filtered]
    median_h = float(np.median(heights)) if heights else 10.0
    row_tol = max(3.5, median_h * 0.55)

    # 1. Group words into visual lines by vertical center
    sorted_by_y = sorted(filtered, key=lambda w: ((w[1] + w[3]) / 2.0, w[0]))
    visual_lines = []
    for w in sorted_by_y:
        cy = (w[1] + w[3]) / 2.0
        if not visual_lines:
            visual_lines.append({"cy": cy, "words": [w]})
        else:
            if abs(cy - visual_lines[-1]["cy"]) <= row_tol:
                visual_lines[-1]["words"].append(w)
                # Update running average cy
                all_cys = [( item[1] + item[3] ) / 2.0 for item in visual_lines[-1]["words"]]
                visual_lines[-1]["cy"] = sum(all_cys) / len(all_cys)
            else:
                visual_lines.append({"cy": cy, "words": [w]})

    # 2. Within each visual line, merge adjacent words that belong to the same cell/phrase
    # Words in the same phrase have a small horizontal gap; words in different table columns have a larger gap
    cell_gap_threshold = max(median_h * 0.95, page_width * 0.018, 9.0)
    line_segments = []
    for vline in visual_lines:
        lwords = sorted(vline["words"], key=lambda w: w[0])
        segs = []
        for w in lwords:
            wx0, wy0, wx1, wy1, wt = w[0], w[1], w[2], w[3], str(w[4]).strip()
            if not segs:
                segs.append({"x0": wx0, "y0": wy0, "x1": wx1, "y1": wy1, "text": wt})
            else:
                gap = wx0 - segs[-1]["x1"]
                if gap <= cell_gap_threshold:
                    segs[-1]["x1"] = max(segs[-1]["x1"], wx1)
                    segs[-1]["y0"] = min(segs[-1]["y0"], wy0)
                    segs[-1]["y1"] = max(segs[-1]["y1"], wy1)
                    segs[-1]["text"] += " " + wt
                else:
                    segs.append({"x0": wx0, "y0": wy0, "x1": wx1, "y1": wy1, "text": wt})
        line_segments.append({
            "y0": min(s["y0"] for s in segs),
            "y1": max(s["y1"] for s in segs),
            "cy": vline["cy"],
            "segments": segs,
        })

    # 3. Identify contiguous groups of lines that have >= 2 columns (allowing 1 single-segment line inside if close)
    table_candidates = []
    curr_group = []
    for idx, lseg in enumerate(line_segments):
        is_multi = len(lseg["segments"]) >= 2
        if is_multi:
            if curr_group:
                prev_y1 = curr_group[-1]["y1"]
                if (lseg["y0"] - prev_y1) > median_h * 4.5:
                    if len(curr_group) >= 2:
                        table_candidates.append(curr_group)
                    curr_group = [lseg]
                else:
                    curr_group.append(lseg)
            else:
                curr_group = [lseg]
        else:
            if len(curr_group) >= 2:
                table_candidates.append(curr_group)
            curr_group = []

    if len(curr_group) >= 2:
        table_candidates.append(curr_group)

    detected_tables = []
    for group in table_candidates:
        # Determine column anchors from segment x0 positions
        all_x0s = [s["x0"] for line in group for s in line["segments"]]
        col_anchors = cluster_1d_positions(all_x0s, tolerance=max(cell_gap_threshold * 1.4, page_width * 0.04))
        if len(col_anchors) < 2:
            continue

        num_cols = len(col_anchors)
        rows_matrix = []
        for line in group:
            row_cells = [""] * num_cols
            for seg in line["segments"]:
                # Assign segment to nearest column anchor
                best_c = min(range(num_cols), key=lambda c: abs(seg["x0"] - col_anchors[c]))
                if row_cells[best_c]:
                    row_cells[best_c] += " " + seg["text"]
                else:
                    row_cells[best_c] = seg["text"]
            rows_matrix.append(row_cells)

        cleaned = clean_table_rows(rows_matrix)
        if len(cleaned) >= 2 and len(cleaned[0]) >= 2:
            tx0 = min(s["x0"] for line in group for s in line["segments"])
            ty0 = min(line["y0"] for line in group)
            tx1 = max(s["x1"] for line in group for s in line["segments"])
            ty1 = max(line["y1"] for line in group)
            detected_tables.append({
                "bbox": (tx0, ty0, tx1, ty1),
                "rows": cleaned,
            })

    return detected_tables


def words_to_text_blocks(words, exclude_bboxes=None):
    """
    Groups words (outside of exclude_bboxes) into logical lines/paragraphs for full-file conversion.
    """
    exclude_bboxes = exclude_bboxes or []

    def is_in_table(wx0, wy0, wx1, wy1):
        cx = (wx0 + wx1) / 2.0
        cy = (wy0 + wy1) / 2.0
        for bx0, by0, bx1, by1 in exclude_bboxes:
            if bx0 - 3 <= cx <= bx1 + 3 and by0 - 3 <= cy <= by1 + 3:
                return True
        return False

    rem = [w for w in words if not is_in_table(w[0], w[1], w[2], w[3]) and str(w[4]).strip()]
    if not rem:
        return []

    heights = [max(2.0, w[3] - w[1]) for w in rem]
    median_h = float(np.median(heights)) if heights else 11.0
    row_tol = max(3.5, median_h * 0.55)

    sorted_w = sorted(rem, key=lambda w: ((w[1] + w[3]) / 2.0, w[0]))
    lines = []
    for w in sorted_w:
        cy = (w[1] + w[3]) / 2.0
        if not lines or abs(cy - lines[-1]["cy"]) > row_tol:
            lines.append({"cy": cy, "words": [w]})
        else:
            lines[-1]["words"].append(w)
            cys = [(item[1] + item[3]) / 2.0 for item in lines[-1]["words"]]
            lines[-1]["cy"] = sum(cys) / len(cys)

    blocks = []
    for line in lines:
        lwords = sorted(line["words"], key=lambda w: w[0])
        text = " ".join(str(w[4]).strip() for w in lwords)
        x0 = min(w[0] for w in lwords)
        y0 = min(w[1] for w in lwords)
        x1 = max(w[2] for w in lwords)
        y1 = max(w[3] for w in lwords)
        avg_h = sum((w[3] - w[1]) for w in lwords) / len(lwords)
        blocks.append({
            "bbox": (x0, y0, x1, y1),
            "text": text,
            "font_size": max(9.0, min(24.0, avg_h * 0.85)),
            "is_bold": avg_h > (median_h * 1.25),
        })
    return blocks


# =====================================================================
# 3. UNIFIED DOCUMENT ANALYZER (PDF & IMAGES)
# =====================================================================
def analyze_source_file(input_path: str, ext: str):
    """
    Analyzes either a PDF or an Image file and returns a list of page dictionaries:
    [
        {
            "page_num": int,
            "width": float,
            "height": float,
            "tables": [{"bbox": (x0, y0, x1, y1), "rows": [[...]]}],
            "text_blocks": [{"bbox": (x0, y0, x1, y1), "text": str, "font_size": float, "is_bold": bool}],
            "images": [{"bbox": (x0, y0, x1, y1), "bytes": bytes}],
            "page_png": bytes,
        }
    ]
    """
    pages_data = []

    if ext in IMAGE_EXTENSIONS:
        with Image.open(input_path) as im:
            im_rgb = im.convert("RGB")
            w, h = im_rgb.size
            buf = io.BytesIO()
            im_rgb.save(buf, format="PNG")
            page_png = buf.getvalue()

        ocr_words = run_windows_ocr(input_path)
        tables = detect_grid_tables_in_image(input_path, ocr_words)
        existing_bboxes = [t["bbox"] for t in tables]
        aligned_tables = extract_tables_from_words(ocr_words, w, h, exclude_bboxes=existing_bboxes)
        tables.extend(aligned_tables)

        all_table_bboxes = [t["bbox"] for t in tables]
        text_blocks = words_to_text_blocks(ocr_words, exclude_bboxes=all_table_bboxes)

        # If the image has neither OCR text nor tables (e.g. a photo/diagram), include the image itself
        embedded_images = []
        if not text_blocks and not tables:
            embedded_images.append({"bbox": (0, 0, w, h), "bytes": page_png})

        pages_data.append({
            "page_num": 1,
            "width": float(w),
            "height": float(h),
            "tables": tables,
            "text_blocks": text_blocks,
            "images": embedded_images,
            "page_png": page_png,
        })
        return pages_data

    # PDF Processing
    doc = fitz.open(input_path)
    for page_idx in range(len(doc)):
        page = doc[page_idx]
        pw = float(page.rect.width)
        ph = float(page.rect.height)

        pix = page.get_pixmap(dpi=150)
        page_png = pix.tobytes("png")

        tables = []
        # 1. Try PyMuPDF native table detection (lines strategy first)
        try:
            found_tabs = page.find_tables(strategy="lines")
            for tab in found_tabs.tables:
                cleaned = clean_table_rows(tab.extract())
                if len(cleaned) >= 2 and len(cleaned[0]) >= 2:
                    tables.append({
                        "bbox": tuple(tab.bbox),
                        "rows": cleaned,
                    })
        except Exception as e:
            print(f"PyMuPDF find_tables error on page {page_idx + 1}: {e}")

        raw_words = page.get_text("words")  # (x0, y0, x1, y1, "word", block_no, line_no, word_no)
        words = [(w[0], w[1], w[2], w[3], w[4]) for w in raw_words]

        if words:
            # Also detect borderless/text-aligned tables outside already-detected bordered tables
            existing_bboxes = [t["bbox"] for t in tables]
            aligned_tables = extract_tables_from_words(words, pw, ph, exclude_bboxes=existing_bboxes)
            tables.extend(aligned_tables)

            all_table_bboxes = [t["bbox"] for t in tables]
            # Extract text blocks outside tables using PyMuPDF dict blocks for accurate font sizes
            text_blocks = []
            page_dict = page.get_text("dict")
            for blk in page_dict.get("blocks", []):
                if blk.get("type") == 0:  # text block
                    bx0, by0, bx1, by1 = blk.get("bbox", (0, 0, 0, 0))
                    # Check if block is inside a detected table
                    cx = (bx0 + bx1) / 2.0
                    cy = (by0 + by1) / 2.0
                    in_tab = any(
                        tx0 - 4 <= cx <= tx1 + 4 and ty0 - 4 <= cy <= ty1 + 4
                        for (tx0, ty0, tx1, ty1) in all_table_bboxes
                    )
                    if in_tab:
                        continue

                    lines_str = []
                    max_size = 11.0
                    is_bold = False
                    for line in blk.get("lines", []):
                        spans_text = []
                        for span in line.get("spans", []):
                            st = span.get("text", "").strip()
                            if st:
                                spans_text.append(st)
                                sz = float(span.get("size", 11.0))
                                if sz > max_size:
                                    max_size = sz
                                font_flags = int(span.get("flags", 0))
                                font_name = span.get("font", "").lower()
                                if (font_flags & 16) or "bold" in font_name or "-bo" in font_name:
                                    is_bold = True
                        if spans_text:
                            lines_str.append(" ".join(spans_text))
                    full_text = "\n".join(lines_str).strip()
                    if full_text:
                        text_blocks.append({
                            "bbox": (bx0, by0, bx1, by1),
                            "text": full_text,
                            "font_size": max_size,
                            "is_bold": is_bold,
                        })
        else:
            # Scanned PDF page (no embedded text) -> Run Windows OCR & image table detection!
            fd_scan, scan_png_path = tempfile.mkstemp(suffix=".png")
            with os.fdopen(fd_scan, "wb") as f_scan:
                f_scan.write(page_png)
            try:
                ocr_words = run_windows_ocr(scan_png_path)
                grid_tabs = detect_grid_tables_in_image(scan_png_path, ocr_words)
                tables.extend(grid_tabs)
                existing_bboxes = [t["bbox"] for t in tables]
                aligned_tabs = extract_tables_from_words(ocr_words, pix.width, pix.height, exclude_bboxes=existing_bboxes)
                tables.extend(aligned_tabs)
                all_table_bboxes = [t["bbox"] for t in tables]
                text_blocks = words_to_text_blocks(ocr_words, exclude_bboxes=all_table_bboxes)
            finally:
                cleanup_files([scan_png_path])

        # Extract embedded images from PDF page
        embedded_images = []
        try:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                rects = page.get_image_rects(xref)
                base_img = doc.extract_image(xref)
                if base_img and base_img.get("image"):
                    img_bytes = base_img["image"]
                    # Convert to PNG via PIL to ensure compatibility with Word/Excel/PPTX
                    with Image.open(io.BytesIO(img_bytes)) as pil_im:
                        out_b = io.BytesIO()
                        pil_im.convert("RGB").save(out_b, format="PNG")
                        png_bytes = out_b.getvalue()
                    bbox = tuple(rects[0]) if rects else (0.0, 0.0, 200.0, 150.0)
                    embedded_images.append({"bbox": bbox, "bytes": png_bytes})
        except Exception as e:
            print(f"Warning extracting images on page {page_idx + 1}: {e}")

        pages_data.append({
            "page_num": page_idx + 1,
            "width": pw,
            "height": ph,
            "tables": tables,
            "text_blocks": text_blocks,
            "images": embedded_images,
            "page_png": page_png,
        })

    doc.close()
    return pages_data


# =====================================================================
# 4. GENERATORS FOR EXCEL (.XLSX), WORD (.DOCX), POWERPOINT (.PPTX)
# =====================================================================
def try_parse_number(val: str):
    """Attempts to parse a clean numeric string so Excel stores numbers as numbers."""
    if not isinstance(val, str):
        return val
    s = val.strip()
    if not s:
        return ""
    # Check standard integer or float (avoiding phone numbers or leading-zero codes like '0012')
    if len(s) > 1 and s.startswith("0") and not s.startswith("0."):
        return s
    clean = s.replace("$", "").replace(",", "").strip()
    if re.fullmatch(r"-?\d+(\.\d+)?", clean):
        try:
            return float(clean) if "." in clean else int(clean)
        except Exception:
            return s
    return s


def generate_excel(pages_data, output_path: str, mode: str = "all"):
    """
    Generates an Excel (.xlsx) file from analyzed pages_data.
    mode: 'tables' (only detected tables) or 'all' (full document content + tables).
    """
    all_tables = []
    for p in pages_data:
        for t in p["tables"]:
            all_tables.append((p["page_num"], t["rows"]))

    if mode == "tables" and not all_tables:
        raise HTTPException(
            status_code=400,
            detail="No se detectaron tablas en el archivo seleccionado. Si deseas convertir todo el contenido, elige la opción 'Convertir todo el archivo'.",
        )

    workbook = xlsxwriter.Workbook(output_path)

    # Formats
    title_fmt = workbook.add_format({
        "bold": True,
        "font_size": 14,
        "font_color": "#1E293B",
        "bottom": 2,
        "bottom_color": "#6366F1",
    })
    subtitle_fmt = workbook.add_format({
        "bold": True,
        "font_size": 11,
        "font_color": "#4F46E5",
    })
    header_fmt = workbook.add_format({
        "bold": True,
        "font_color": "#FFFFFF",
        "bg_color": "#4F46E5",
        "border": 1,
        "border_color": "#CBD5E1",
        "valign": "vcenter",
        "text_wrap": True,
    })
    cell_fmt = workbook.add_format({
        "border": 1,
        "border_color": "#CBD5E1",
        "valign": "vcenter",
        "text_wrap": True,
    })
    cell_alt_fmt = workbook.add_format({
        "bg_color": "#F8FAFC",
        "border": 1,
        "border_color": "#CBD5E1",
        "valign": "vcenter",
        "text_wrap": True,
    })
    text_fmt = workbook.add_format({
        "font_size": 11,
        "valign": "top",
        "text_wrap": True,
    })
    bold_text_fmt = workbook.add_format({
        "bold": True,
        "font_size": 12,
        "font_color": "#0F172A",
        "valign": "top",
    })

    def write_table_to_sheet(ws, start_row: int, rows):
        col_widths = {}
        for r_idx, row in enumerate(rows):
            fmt = header_fmt if r_idx == 0 else (cell_alt_fmt if r_idx % 2 == 1 else cell_fmt)
            for c_idx, val in enumerate(row):
                parsed_val = val if r_idx == 0 else try_parse_number(val)
                ws.write(start_row + r_idx, c_idx, parsed_val, fmt)
                str_len = len(str(val or ""))
                col_widths[c_idx] = max(col_widths.get(c_idx, 10), min(50, str_len + 4))
        return start_row + len(rows), col_widths

    if mode == "tables":
        # Create one dedicated sheet per detected table
        for idx, (page_num, rows) in enumerate(all_tables):
            sheet_name = f"Tabla_{idx + 1}_Pag_{page_num}"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.write(0, 0, f"Tabla {idx + 1} (Extraída de Página {page_num})", subtitle_fmt)
            _, col_widths = write_table_to_sheet(ws, 2, rows)
            for c_idx, width in col_widths.items():
                ws.set_column(c_idx, c_idx, width)
    else:
        # Mode == "all": Full document per page (text + tables + images in reading order)
        for p in pages_data:
            sheet_name = f"Pagina_{p['page_num']}"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.write(0, 0, f"Contenido de Página {p['page_num']}", title_fmt)

            # Combine text blocks and tables sorted by vertical position y0
            elements = []
            for tb in p["text_blocks"]:
                elements.append({"type": "text", "y0": tb["bbox"][1], "data": tb})
            for tab in p["tables"]:
                elements.append({"type": "table", "y0": tab["bbox"][1], "data": tab})
            elements.sort(key=lambda el: el["y0"])

            curr_row = 2
            max_col_widths = {0: 45}

            if not elements and p["images"]:
                # Visual/image-only page: insert the image into the sheet
                img_stream = io.BytesIO(p["images"][0]["bytes"])
                ws.insert_image(curr_row, 0, f"imagen_pag_{p['page_num']}.png", {"image_data": img_stream, "x_scale": 0.5, "y_scale": 0.5})
            else:
                for el in elements:
                    if el["type"] == "text":
                        tb = el["data"]
                        fmt = bold_text_fmt if tb.get("is_bold") else text_fmt
                        for line in tb["text"].split("\n"):
                            if line.strip():
                                ws.write(curr_row, 0, line.strip(), fmt)
                                max_col_widths[0] = max(max_col_widths.get(0, 30), min(80, len(line.strip()) + 4))
                                curr_row += 1
                        curr_row += 1
                    elif el["type"] == "table":
                        rows = el["data"]["rows"]
                        curr_row, col_w = write_table_to_sheet(ws, curr_row, rows)
                        for c_idx, w_val in col_w.items():
                            max_col_widths[c_idx] = max(max_col_widths.get(c_idx, 12), w_val)
                        curr_row += 2

            for c_idx, width in max_col_widths.items():
                ws.set_column(c_idx, c_idx, width)

        # Also add individual sheets for each detected table for convenience
        for idx, (page_num, rows) in enumerate(all_tables):
            sheet_name = f"Tabla_{idx + 1}"[:31]
            ws = workbook.add_worksheet(sheet_name)
            ws.write(0, 0, f"Tabla {idx + 1} (Página {page_num})", subtitle_fmt)
            _, col_widths = write_table_to_sheet(ws, 2, rows)
            for c_idx, width in col_widths.items():
                ws.set_column(c_idx, c_idx, width)

    workbook.close()


def set_docx_cell_background(cell, fill_hex: str):
    """Sets background color of a python-docx table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def add_table_to_docx(doc: DocxDocument, rows):
    """Adds a formatted editable table to a Word document."""
    if not rows or not rows[0]:
        return
    num_rows = len(rows)
    num_cols = len(rows[0])
    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.style = "Table Grid"

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if c_idx >= num_cols:
                continue
            cell = table.cell(r_idx, c_idx)
            cell.text = str(val or "")
            if r_idx == 0:
                set_docx_cell_background(cell, "4F46E5")
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
                        run.font.color.rgb = DocxRGBColor(255, 255, 255)
                        run.font.size = DocxPt(10)
            else:
                if r_idx % 2 == 1:
                    set_docx_cell_background(cell, "F8FAFC")
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = DocxPt(10)


def generate_word(pages_data, output_path: str, mode: str = "all"):
    """
    Generates a Word (.docx) file from analyzed pages_data.
    mode: 'tables' (only detected tables) or 'all' (full document content + tables + images).
    """
    all_tables = []
    for p in pages_data:
        for t in p["tables"]:
            all_tables.append((p["page_num"], t["rows"]))

    if mode == "tables" and not all_tables:
        raise HTTPException(
            status_code=400,
            detail="No se detectaron tablas en el archivo seleccionado. Si deseas convertir todo el contenido, elige la opción 'Convertir todo el archivo'.",
        )

    doc = DocxDocument()

    if mode == "tables":
        heading = doc.add_heading("Tablas Detectadas", level=1)
        for idx, (page_num, rows) in enumerate(all_tables):
            sub = doc.add_heading(f"Tabla {idx + 1} (Página {page_num})", level=2)
            add_table_to_docx(doc, rows)
            doc.add_paragraph("")
    else:
        for p_idx, p in enumerate(pages_data):
            if p_idx > 0:
                doc.add_page_break()

            elements = []
            for tb in p["text_blocks"]:
                elements.append({"type": "text", "y0": tb["bbox"][1], "data": tb})
            for tab in p["tables"]:
                elements.append({"type": "table", "y0": tab["bbox"][1], "data": tab})
            for img in p["images"]:
                elements.append({"type": "image", "y0": img["bbox"][1], "data": img})

            elements.sort(key=lambda el: el["y0"])

            if not elements:
                # Fallback: embed full page image if no text/tables/images were extracted
                img_stream = io.BytesIO(p["page_png"])
                doc.add_picture(img_stream, width=DocxInches(6.0))
                continue

            for el in elements:
                if el["type"] == "text":
                    tb = el["data"]
                    para = doc.add_paragraph()
                    run = para.add_run(tb["text"])
                    run.bold = bool(tb.get("is_bold"))
                    sz = float(tb.get("font_size", 11.0))
                    run.font.size = DocxPt(max(9.0, min(22.0, sz)))
                elif el["type"] == "table":
                    add_table_to_docx(doc, el["data"]["rows"])
                    doc.add_paragraph("")
                elif el["type"] == "image":
                    try:
                        img_stream = io.BytesIO(el["data"]["bytes"])
                        doc.add_picture(img_stream, width=DocxInches(4.5))
                    except Exception as e:
                        print(f"Could not insert image into Word: {e}")

    doc.save(output_path)


def add_editable_table_slide(prs: Presentation, title_text: str, rows):
    """Adds a slide with a native editable PowerPoint table."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Title box
    title_box = slide.shapes.add_textbox(PptxInches(0.6), PptxInches(0.4), PptxInches(12.0), PptxInches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = PptxPt(20)
    p.font.bold = True
    p.font.color.rgb = PptxRGBColor(30, 41, 59)

    num_rows = len(rows)
    num_cols = len(rows[0]) if rows else 1

    table_w = PptxInches(12.1)
    table_h = PptxInches(min(5.8, max(1.5, num_rows * 0.42)))
    graphic_frame = slide.shapes.add_table(
        num_rows, num_cols, PptxInches(0.6), PptxInches(1.2), table_w, table_h
    )
    table = graphic_frame.table

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            if c_idx >= num_cols:
                continue
            cell = table.cell(r_idx, c_idx)
            cell.text = str(val or "")
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = PptxRGBColor(79, 70, 229)
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.bold = True
                        run.font.size = PptxPt(11)
                        run.font.color.rgb = PptxRGBColor(255, 255, 255)
            else:
                cell.fill.solid()
                if r_idx % 2 == 1:
                    cell.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
                else:
                    cell.fill.fore_color.rgb = PptxRGBColor(255, 255, 255)
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = PptxPt(10)
                        run.font.color.rgb = PptxRGBColor(15, 23, 42)


def generate_pptx(input_path: str, ext: str, pages_data, output_path: str, mode: str = "all"):
    """
    Generates a PowerPoint (.pptx) presentation.
    mode: 'tables' (creates slides with native editable PowerPoint tables for each detected table)
          'all' (converts entire file into slides, preserving visual layout and auto-splitting vertical dual-slide PDFs, plus editable table slides if tables exist).
    """
    all_tables = []
    for p in pages_data:
        for t in p["tables"]:
            all_tables.append((p["page_num"], t["rows"]))

    if mode == "tables" and not all_tables:
        raise HTTPException(
            status_code=400,
            detail="No se detectaron tablas en el archivo seleccionado. Si deseas convertir todo el contenido, elige la opción 'Convertir todo el archivo'.",
        )

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_slide_layout = prs.slide_layouts[6]

    if mode == "tables":
        # Chunk large tables into max 14 rows per slide (including header)
        MAX_ROWS_PER_SLIDE = 14
        for idx, (page_num, rows) in enumerate(all_tables):
            header = rows[0]
            data_rows = rows[1:]
            if len(rows) <= MAX_ROWS_PER_SLIDE or not data_rows:
                add_editable_table_slide(prs, f"Tabla {idx + 1} (Página {page_num})", rows)
            else:
                for chunk_idx in range(0, len(data_rows), MAX_ROWS_PER_SLIDE - 1):
                    chunk = [header] + data_rows[chunk_idx:chunk_idx + MAX_ROWS_PER_SLIDE - 1]
                    part_num = (chunk_idx // (MAX_ROWS_PER_SLIDE - 1)) + 1
                    add_editable_table_slide(
                        prs,
                        f"Tabla {idx + 1} - Parte {part_num} (Página {page_num})",
                        chunk,
                    )
        prs.save(output_path)
        return

    # Mode == "all": Full document to PowerPoint
    if ext == "pdf":
        doc = fitz.open(input_path)
        for page_num in range(len(doc)):
            page = doc[page_num]
            w = page.rect.width
            h = page.rect.height

            # Heurística de división: Si es Vertical (Portrait), asumimos 2 diapositivas por hoja (mitad superior e inferior)
            if h > w and (h / w) > 1.1:
                rect1 = fitz.Rect(0, 0, w, h / 2)
                rect2 = fitz.Rect(0, h / 2, w, h)
                for rect in [rect1, rect2]:
                    pix = page.get_pixmap(clip=rect, dpi=150)
                    img_stream = io.BytesIO(pix.tobytes("png"))
                    slide = prs.slides.add_slide(blank_slide_layout)
                    slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            else:
                pix = page.get_pixmap(dpi=150)
                img_stream = io.BytesIO(pix.tobytes("png"))
                slide = prs.slides.add_slide(blank_slide_layout)
                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
        doc.close()
    else:
        # Image input -> Add image slide maintaining aspect ratio centered on slide
        for p in pages_data:
            slide = prs.slides.add_slide(blank_slide_layout)
            img_stream = io.BytesIO(p["page_png"])
            slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Also append editable table slides at the end if tables were detected in the file
    if all_tables:
        MAX_ROWS_PER_SLIDE = 14
        for idx, (page_num, rows) in enumerate(all_tables):
            header = rows[0]
            data_rows = rows[1:]
            if len(rows) <= MAX_ROWS_PER_SLIDE or not data_rows:
                add_editable_table_slide(prs, f"Tabla Editable {idx + 1} (Página {page_num})", rows)
            else:
                for chunk_idx in range(0, len(data_rows), MAX_ROWS_PER_SLIDE - 1):
                    chunk = [header] + data_rows[chunk_idx:chunk_idx + MAX_ROWS_PER_SLIDE - 1]
                    part_num = (chunk_idx // (MAX_ROWS_PER_SLIDE - 1)) + 1
                    add_editable_table_slide(
                        prs,
                        f"Tabla Editable {idx + 1} - Parte {part_num} (Página {page_num})",
                        chunk,
                    )

    prs.save(output_path)


def image_to_pdf(input_path: str, output_path: str):
    """Converts an image file into a single-page PDF."""
    with Image.open(input_path) as im:
        im_rgb = im.convert("RGB")
        im_rgb.save(output_path, "PDF", resolution=150.0)


def office_to_pdf(input_path: str, output_path: str, ext: str):
    """Converts DOCX or PPTX to PDF using COM if available, or PyMuPDF fallback."""
    abs_in = os.path.abspath(input_path)
    abs_out = os.path.abspath(output_path)
    try:
        import comtypes.client
        comtypes.CoInitialize()
        if ext == "pptx":
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            try:
                deck = powerpoint.Presentations.Open(abs_in, WithWindow=False)
                deck.SaveAs(abs_out, 32)  # 32 = ppSaveAsPDF
                deck.Close()
            finally:
                powerpoint.Quit()
            return
        elif ext == "docx":
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(abs_in)
                doc.SaveAs(abs_out, FileFormat=17)  # 17 = wdFormatPDF
                doc.Close()
            finally:
                word.Quit()
            return
    except Exception as e:
        print(f"Office COM fallback triggered ({e}), generating PDF via Python...")

    # Fallback if MS Office is not installed
    pdf_doc = fitz.open()
    if ext == "docx":
        d = DocxDocument(input_path)
        page = pdf_doc.new_page()
        y = 50
        for para in d.paragraphs:
            txt = para.text.strip()
            if not txt:
                continue
            if y > page.rect.height - 60:
                page = pdf_doc.new_page()
                y = 50
            page.insert_text(fitz.Point(50, y), txt[:110], fontsize=11)
            y += 18
    elif ext == "pptx":
        prs = Presentation(input_path)
        for slide in prs.slides:
            page = pdf_doc.new_page(width=842, height=595)
            y = 60
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for p in shape.text_frame.paragraphs:
                        txt = p.text.strip()
                        if txt:
                            page.insert_text(fitz.Point(50, y), txt[:110], fontsize=12)
                            y += 20
    if len(pdf_doc) == 0:
        pdf_doc.new_page()
    pdf_doc.save(output_path)
    pdf_doc.close()


def cleanup_files(paths):
    for p in paths:
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass


# =====================================================================
# 5. CONVERTER ENDPOINT (/api/convert)
# =====================================================================
@app.post("/api/convert")
async def convert_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    target_format: str = Form(...),
    conversion_mode: str = Form("all"),  # "all" (todo el archivo) or "tables" (solo tablas)
):
    if not file.filename or "." not in file.filename:
        raise HTTPException(status_code=400, detail="El archivo no tiene una extensión válida.")

    ext = file.filename.rsplit(".", 1)[-1].lower()
    target_format = target_format.lower().strip()
    mode = conversion_mode.lower().strip()
    if mode not in ("all", "tables"):
        mode = "all"

    fd_in, in_path = tempfile.mkstemp(suffix=f".{ext}")
    with os.fdopen(fd_in, "wb") as f:
        f.write(await file.read())

    fd_out, out_path = tempfile.mkstemp(suffix=f".{target_format}")
    os.close(fd_out)

    MEDIA_TYPES = {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    try:
        if ext == "pdf" or ext in IMAGE_EXTENSIONS:
            if target_format in ("xlsx", "docx", "pptx"):
                pages_data = analyze_source_file(in_path, ext)
                if target_format == "xlsx":
                    generate_excel(pages_data, out_path, mode=mode)
                elif target_format == "docx":
                    generate_word(pages_data, out_path, mode=mode)
                elif target_format == "pptx":
                    generate_pptx(in_path, ext, pages_data, out_path, mode=mode)
            elif ext in IMAGE_EXTENSIONS and target_format == "pdf":
                image_to_pdf(in_path, out_path)
            else:
                raise HTTPException(status_code=400, detail=f"Conversión de .{ext} a .{target_format} no soportada.")

        elif ext in ("docx", "pptx") and target_format == "pdf":
            office_to_pdf(in_path, out_path, ext)

        else:
            raise HTTPException(status_code=400, detail=f"Conversión de .{ext} a .{target_format} no soportada.")

        media_type = MEDIA_TYPES.get(target_format, "application/octet-stream")
        background_tasks.add_task(cleanup_files, [in_path, out_path])
        return FileResponse(out_path, media_type=media_type, filename=f"convertido.{target_format}")

    except HTTPException:
        cleanup_files([in_path, out_path])
        raise
    except Exception as e:
        tb = traceback.format_exc()
        print(f"\n=== ERROR EN /api/convert ===\n{tb}\n=============================")
        cleanup_files([in_path, out_path])
        raise HTTPException(status_code=500, detail=str(e))
