import os
import json
import re
import cv2
import numpy as np
from flask import Flask, request, jsonify, send_file, render_template
from pyzbar.pyzbar import decode as pyzbar_decode
from PIL import Image
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import tempfile

# EasyOCR loaded lazily on first use (heavy import ~1GB model download on first run)
_ocr_reader = None

def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(["en"], gpu=False, verbose=False)
    return _ocr_reader

# Minimum digit length to be considered a valid SKU
SKU_MIN_DIGITS = 4

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Store inventory data in a file so it persists across requests
DATA_FILE = os.environ.get("DATA_FILE", "inventory_data.json")


def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _barcode_scan(frame):
    """Run pyzbar on a frame, return set of decoded strings."""
    found = set()
    pil_img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
    for r in pyzbar_decode(pil_img):
        code = r.data.decode("utf-8").strip()
        if code:
            found.add(code)
    return found


def _ocr_scan(frame):
    """Run EasyOCR on a frame, return set of digit sequences that look like SKUs."""
    found = set()
    try:
        reader = get_ocr_reader()
        results = reader.readtext(frame, detail=1, paragraph=False)
        for (_bbox, text, conf) in results:
            if conf < 0.5:
                continue
            # Keep only runs of digits (with optional dashes/slashes inside)
            for token in text.split():
                clean = re.sub(r"[^0-9\-\/]", "", token)
                digits_only = re.sub(r"[^0-9]", "", clean)
                if len(digits_only) >= SKU_MIN_DIGITS:
                    found.add(clean.strip("-/"))
    except Exception:
        pass
    return found


def extract_codes_from_frame(frame, use_ocr=True):
    """Try barcode first; fall back to OCR if nothing found."""
    codes = _barcode_scan(frame)

    # Enhance contrast and retry barcode if empty
    if not codes:
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        codes = _barcode_scan(cv2.cvtColor(enhanced, cv2.COLOR_GRAY2BGR))

    # OCR fallback
    if not codes and use_ocr:
        codes = _ocr_scan(frame)

    return codes


def process_video(file_bytes, sample_every=15):
    """Extract codes from video bytes. Returns dict of {code: count}."""
    counts = {}

    with tempfile.NamedTemporaryFile(delete=False, suffix=".mp4") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        cap = cv2.VideoCapture(tmp_path)
        frame_num = 0
        seen_in_sequence = {}

        while True:
            ret, frame = cap.read()
            if not ret:
                break

            if frame_num % sample_every == 0:
                found = extract_codes_from_frame(frame)
                for code in found:
                    last_seen = seen_in_sequence.get(code, -999)
                    if frame_num - last_seen > 30 * sample_every:
                        counts[code] = counts.get(code, 0) + 1
                    seen_in_sequence[code] = frame_num

            frame_num += 1

        cap.release()
    finally:
        os.unlink(tmp_path)

    return counts


def process_image(file_bytes):
    """Extract codes from image bytes. Returns dict of {code: 1}."""
    nparr = np.frombuffer(file_bytes, np.uint8)
    frame = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if frame is None:
        return {}
    found = extract_codes_from_frame(frame)
    return {code: 1 for code in found}


def build_excel(inventory_data):
    """Build Excel workbook from inventory_data: {shelf: {sku: count}}."""
    wb = openpyxl.Workbook()

    # ── Summary sheet ──────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "סיכום"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F497D")
    alt_fill = PatternFill("solid", start_color="DCE6F1")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Collect all SKUs
    all_skus = sorted({sku for shelf_data in inventory_data.values() for sku in shelf_data})
    shelves = sorted(inventory_data.keys())

    # Headers row
    ws_sum.append(["מקט / מדף"] + shelves + ["סה\"כ"])
    for col, cell in enumerate(ws_sum[1], 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    ws_sum.row_dimensions[1].height = 25

    for row_idx, sku in enumerate(all_skus, 2):
        row = [sku]
        for shelf in shelves:
            row.append(inventory_data.get(shelf, {}).get(sku, 0))
        # Total formula
        first_col = get_column_letter(2)
        last_col = get_column_letter(1 + len(shelves))
        row.append(f"=SUM({first_col}{row_idx}:{last_col}{row_idx})")
        ws_sum.append(row)

        fill = alt_fill if row_idx % 2 == 0 else None
        for col_idx, cell in enumerate(ws_sum[row_idx], 1):
            cell.font = Font(name="Arial", size=10)
            cell.alignment = center
            cell.border = border
            if fill:
                cell.fill = fill

    # Totals row
    total_row = ws_sum.max_row + 1
    ws_sum.cell(total_row, 1, 'סה"כ')
    ws_sum.cell(total_row, 1).font = Font(name="Arial", bold=True, size=10)
    ws_sum.cell(total_row, 1).fill = PatternFill("solid", start_color="BDD7EE")
    for col_idx in range(2, 2 + len(shelves) + 1):
        col_letter = get_column_letter(col_idx)
        cell = ws_sum.cell(total_row, col_idx,
                           f"=SUM({col_letter}2:{col_letter}{total_row - 1})")
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.fill = PatternFill("solid", start_color="BDD7EE")
        cell.alignment = center
        cell.border = border

    ws_sum.column_dimensions["A"].width = 22
    for col_idx in range(2, 2 + len(shelves) + 1):
        ws_sum.column_dimensions[get_column_letter(col_idx)].width = 12

    # ── Per-shelf sheets ───────────────────────────────────────────────
    for shelf in shelves:
        ws = wb.create_sheet(f"מדף {shelf}")
        ws.append(["מקט", "כמות"])
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center
            cell.border = border
        ws.row_dimensions[1].height = 25

        for row_idx, (sku, qty) in enumerate(
            sorted(inventory_data[shelf].items()), 2
        ):
            ws.append([sku, qty])
            for cell in ws[row_idx]:
                cell.font = Font(name="Arial", size=10)
                cell.alignment = center
                cell.border = border
            if row_idx % 2 == 0:
                for cell in ws[row_idx]:
                    cell.fill = alt_fill

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 12

    return wb


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/shelves", methods=["GET"])
def get_shelves():
    data = load_data()
    return jsonify({
        "shelves": list(data.keys()),
        "total_skus": sum(len(v) for v in data.values()),
        "data": data
    })


@app.route("/api/upload", methods=["POST"])
def upload():
    shelf = request.form.get("shelf", "").strip()
    if not shelf:
        return jsonify({"error": "חסר מספר מדף"}), 400

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "לא הועלה קובץ"}), 400

    file_bytes = file.read()
    filename = file.filename.lower()

    if any(filename.endswith(ext) for ext in [".mp4", ".mov", ".avi", ".mkv", ".3gp", ".m4v"]):
        counts = process_video(file_bytes)
    elif any(filename.endswith(ext) for ext in [".jpg", ".jpeg", ".png", ".bmp", ".webp", ".heic"]):
        counts = process_image(file_bytes)
    else:
        return jsonify({"error": "סוג קובץ לא נתמך"}), 400

    if not counts:
        return jsonify({"error": "לא זוהו מקטים בקובץ — נסה תמונה ברורה יותר", "found": 0}), 200

    # Merge into inventory
    data = load_data()
    shelf_data = data.get(shelf, {})
    for sku, qty in counts.items():
        shelf_data[sku] = shelf_data.get(sku, 0) + qty
    data[shelf] = shelf_data
    save_data(data)

    return jsonify({
        "shelf": shelf,
        "found": len(counts),
        "new_items": counts,
        "shelf_total": shelf_data
    })


@app.route("/api/delete_shelf/<shelf>", methods=["DELETE"])
def delete_shelf(shelf):
    data = load_data()
    if shelf in data:
        del data[shelf]
        save_data(data)
    return jsonify({"ok": True})


@app.route("/api/clear", methods=["POST"])
def clear_all():
    save_data({})
    return jsonify({"ok": True})


@app.route("/api/export", methods=["GET"])
def export_excel():
    data = load_data()
    if not data:
        return jsonify({"error": "אין נתונים לייצוא"}), 400

    wb = build_excel(data)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"מלאי_{timestamp}.xlsx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


if __name__ == "__main__":
    os.makedirs("templates", exist_ok=True)
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
