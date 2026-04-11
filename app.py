import io
import os
import re
import uuid
import subprocess
import tempfile

import pdfplumber
from docxtpl import DocxTemplate
from flask import Flask, jsonify, render_template, request, send_file
from pypdf import PdfWriter, PdfReader, Transformation

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit

TEMPLATE_PATH     = os.path.join(os.path.dirname(__file__), "template.docx")
BORDER_ONLY_PATH  = os.path.join(os.path.dirname(__file__), "border_only.pdf")
TMP_DIR = tempfile.gettempdir()


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

DEPARTMENTS = [
    "Department of Accounting and Finance",
    "Department of Supply Chain and Information Systems",
    "Department of Human Resource and Organisational Development",
    "Department of Marketing and Corporate Strategy",
    "Department of Hospitality and Tourism Studies",
]

PROGRAMMES = [
    "BSc Business Administration (Accounting)",
    "BSc Business Administration (Banking & Finance)",
    "BSc Business Administration (Marketing)",
    "BSc Business Administration (International Business)",
    "BSc Business Administration (Human Resource Management)",
    "BSc Business Administration (Management)",
    "BSc Business Administration (Logistics and Supply Chain Management)",
    "BSc Business Administration (Business Information Technology)",
    "BSc Hospitality and Tourism Management",
]


def _match_department(text: str) -> str:
    """Try to match a department from PDF text."""
    text_up = text.upper()
    keywords = {
        "ACCOUNTING AND FINANCE":               DEPARTMENTS[0],
        "SUPPLY CHAIN AND INFORMATION":         DEPARTMENTS[1],
        "HUMAN RESOURCE AND ORGANISATIONAL":    DEPARTMENTS[2],
        "HUMAN RESOURCE":                       DEPARTMENTS[2],
        "MARKETING AND CORPORATE":              DEPARTMENTS[3],
        "MARKETING":                            DEPARTMENTS[3],
        "HOSPITALITY AND TOURISM":              DEPARTMENTS[4],
    }
    for kw, dept in keywords.items():
        if kw in text_up:
            return dept
    return ""


def _match_programme(text: str) -> str:
    """Try to match a programme from PDF text."""
    text_up = text.upper()
    keywords = {
        "BUSINESS INFORMATION TECHNOLOGY":      PROGRAMMES[7],
        "LOGISTICS AND SUPPLY CHAIN":           PROGRAMMES[6],
        "SUPPLY CHAIN":                         PROGRAMMES[6],
        "HUMAN RESOURCE MANAGEMENT":            PROGRAMMES[4],
        "HUMAN RESOURCE":                       PROGRAMMES[4],
        "INTERNATIONAL BUSINESS":               PROGRAMMES[3],
        "BANKING":                              PROGRAMMES[1],
        "MARKETING":                            PROGRAMMES[2],
        "HOSPITALITY AND TOURISM":              PROGRAMMES[8],
        "ACCOUNTING":                           PROGRAMMES[0],
        "MANAGEMENT":                           PROGRAMMES[5],
    }
    for kw, prog in keywords.items():
        if kw in text_up:
            return prog
    return ""


def _name_from_filename(filename: str) -> str:
    """Extract student name from filename pattern: 'NAME LINKEDIN ASSIGNMENT...'"""
    base = os.path.splitext(os.path.basename(filename))[0]
    m = re.match(r"^(.+?)\s+LINKEDIN\s+ASSIGNMENT", base, re.IGNORECASE)
    return m.group(1).strip().title() if m else ""


def extract_pdf_fields(pdf_path: str, original_filename: str = "") -> dict:
    """Extract all available fields from the assignment PDF."""
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

    result = {}

    # Student ID — digits only, stops before INDEX NUMBER on the same line
    m = re.search(r"STUDENT\s+ID\s*:\s*(\d+)", full_text, re.IGNORECASE)
    result["student_id"] = m.group(1).strip() if m else ""

    # Index Number — digits only
    m = re.search(r"INDEX\s+NUMBER\s*:\s*(\d+)", full_text, re.IGNORECASE)
    result["index_number"] = m.group(1).strip() if m else ""

    # Student Name — explicit label first, then fall back to original filename
    m = re.search(r"Student\s+Name\s*[:\-]\s*(.+?)(?:\n|$)", full_text, re.IGNORECASE)
    result["student_name"] = m.group(1).strip() if m else _name_from_filename(original_filename or pdf_path)

    # LinkedIn URL
    m = re.search(r"((?:https?://)?(?:www\.)?linkedin\.com/in/[^\s]+)", full_text, re.IGNORECASE)
    if m:
        url = m.group(1).strip().rstrip(".,;)")
        if not url.startswith("http"):
            url = "https://" + url
        result["linkedin_link"] = url
    else:
        result["linkedin_link"] = ""

    # Department & Programme
    result["department"] = _match_department(full_text)
    result["programme"]  = _match_programme(full_text)

    return result


# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

# A4 dimensions in points (matches the template page size)
_A4_W = 595.28
_A4_H = 841.89

# Inset (pts) from the page edge to safely inside the TableGrid border frame
# matches template margins: left=709dxa, right=707dxa, top=568dxa, bottom=426dxa
_INSET_L = 709 / 1440 * 72   # 35.45 pts
_INSET_R = 707 / 1440 * 72   # 35.35 pts
_INSET_T = 568 / 1440 * 72   # 28.4  pts
_INSET_B = 426 / 1440 * 72   # 21.3  pts
# Add a small padding so content doesn't touch the border line itself
_PAD = 6


def _fit_page_in_border(page) -> None:
    """Scale and centre a page's content to fit inside the A4 border frame."""
    pw = float(page.mediabox.width)
    ph = float(page.mediabox.height)

    avail_w = _A4_W - _INSET_L - _INSET_R - 2 * _PAD
    avail_h = _A4_H - _INSET_T - _INSET_B - 2 * _PAD

    scale = min(avail_w / pw, avail_h / ph)

    sw = pw * scale
    sh = ph * scale

    tx = _INSET_L + _PAD + (avail_w - sw) / 2
    ty = _INSET_B + _PAD + (avail_h - sh) / 2

    page.add_transformation(Transformation().scale(scale).translate(tx, ty))
    page.mediabox.lower_left  = (0, 0)
    page.mediabox.upper_right = (_A4_W, _A4_H)


def _apply_border_to_assignment_pages(cover_bytes: bytes, assignment_bytes: bytes) -> bytes:
    """
    - Cover sheet page: left as-is (table already provides the border frame).
    - Assignment pages: scaled to fit inside the border frame, then border stamped on top.
    """
    writer = PdfWriter()

    # Page 1: cover sheet — no scaling, no extra border overlay
    cover_reader = PdfReader(io.BytesIO(cover_bytes))
    for page in cover_reader.pages:
        writer.add_page(page)

    if not assignment_bytes:
        out = io.BytesIO()
        writer.write(out)
        out.seek(0)
        return out.read()

    border_page = PdfReader(BORDER_ONLY_PATH).pages[0] if os.path.exists(BORDER_ONLY_PATH) else None

    assignment_reader = PdfReader(io.BytesIO(assignment_bytes))
    for page in assignment_reader.pages:
        _fit_page_in_border(page)
        if border_page:
            page.merge_page(border_page, over=True)
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


def _libreoffice_cmd() -> str:
    """Return the correct LibreOffice executable for the current platform."""
    import shutil
    for cmd in ("libreoffice", "soffice"):
        if shutil.which(cmd):
            return cmd
    # Windows fallback: common install path
    win_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    if os.path.exists(win_path):
        return win_path
    raise RuntimeError(
        "LibreOffice is not installed or not found in PATH. "
        "Install it from https://www.libreoffice.org/"
    )


def docx_to_pdf_bytes(docx_path: str) -> bytes:
    """Convert a .docx file to PDF bytes via LibreOffice headless."""
    out_dir = os.path.dirname(docx_path)
    try:
        proc = subprocess.run(
            [
                _libreoffice_cmd(),
                "--headless",
                "--convert-to", "pdf",
                "--outdir", out_dir,
                docx_path,
            ],
            capture_output=True,
            text=True,
            timeout=90,
        )
    except FileNotFoundError:
        raise RuntimeError(
            "LibreOffice is not installed or not found in PATH. "
            "Install it from https://www.libreoffice.org/"
        )

    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice conversion failed:\n{proc.stderr.strip()}")

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise FileNotFoundError("LibreOffice did not produce a PDF output file.")

    with open(pdf_path, "rb") as fh:
        data = fh.read()

    try:
        os.remove(pdf_path)
    except OSError:
        pass

    return data


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/ping")
def ping():
    return {"status": "ok"}


@app.route("/extract", methods=["POST"])
def extract():
    """AJAX endpoint: receive a PDF, return extracted student fields as JSON."""
    pdf_file = request.files.get("pdf_file")
    if not pdf_file or not pdf_file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Please upload a valid PDF file."}), 400

    uid = uuid.uuid4().hex
    # Preserve original filename so name extraction from filename works
    safe_name = re.sub(r"[^\w\s\-.]", "", pdf_file.filename)
    pdf_path = os.path.join(TMP_DIR, f"{uid}_{safe_name}")
    try:
        pdf_file.save(pdf_path)
        data = extract_pdf_fields(pdf_path, original_filename=pdf_file.filename)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        try:
            os.remove(pdf_path)
        except OSError:
            pass

    return jsonify(data)


@app.route("/generate", methods=["POST"])
def generate():
    # --- validate template exists ---
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify(
            {"error": "template.docx not found on the server."}
        ), 500

    # --- read student details from form (entered / auto-filled by student) ---
    student_id   = request.form.get("student_id", "").strip()
    index_number = request.form.get("index_number", "").strip()
    student_name = request.form.get("student_name", "").strip()

    if not student_id:
        return jsonify({"error": "Student ID is required."}), 400
    if not index_number:
        return jsonify({"error": "Index Number is required."}), 400
    if not student_name:
        return jsonify({"error": "Student Name is required."}), 400

    # --- read assignment details ---
    department   = request.form.get("department", "").strip()
    programme    = request.form.get("programme", "").strip()
    linkedin_link = request.form.get("linkedin_link", "").strip()

    if not department:
        return jsonify({"error": "Please select a Department."}), 400
    if not programme:
        return jsonify({"error": "Please select a Programme."}), 400
    if not linkedin_link:
        return jsonify({"error": "LinkedIn URL is required."}), 400

    # --- accept the original assignment PDF for merging ---
    assignment_pdf = request.files.get("pdf_file")

    uid = uuid.uuid4().hex
    docx_output_path = os.path.join(TMP_DIR, f"{uid}_output.docx")
    assignment_pdf_path = None

    try:
        context = {
            "student_id":       student_id,
            "index_number":     index_number,
            "student_name":     student_name,
            "course_code":      request.form.get("course_code", "").strip(),
            "assignment_title": request.form.get("assignment_title", "").strip(),
            "department":       department,
            "programme":        programme,
            "linkedin_link":    linkedin_link,
        }

        # Render cover sheet docx → PDF bytes
        tpl = DocxTemplate(TEMPLATE_PATH)
        tpl.render(context)
        tpl.save(docx_output_path)
        cover_bytes = docx_to_pdf_bytes(docx_output_path)

        # Read assignment PDF bytes (if provided)
        assignment_bytes = None
        if assignment_pdf and assignment_pdf.filename.lower().endswith(".pdf"):
            assignment_pdf_path = os.path.join(TMP_DIR, f"{uid}_assignment.pdf")
            assignment_pdf.save(assignment_pdf_path)
            with open(assignment_pdf_path, "rb") as f:
                assignment_bytes = f.read()

        # Combine: cover sheet as-is + assignment pages scaled & bordered
        pdf_bytes = _apply_border_to_assignment_pages(cover_bytes, assignment_bytes)

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    finally:
        for path in (docx_output_path, assignment_pdf_path):
            if path:
                try:
                    os.remove(path)
                except OSError:
                    pass

    return send_file(
        io.BytesIO(pdf_bytes),
        as_attachment=True,
        download_name="assignment_with_coversheet.pdf",
        mimetype="application/pdf",
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") == "development"
    app.run(debug=debug, host="0.0.0.0", port=port)
