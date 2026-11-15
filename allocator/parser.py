from pathlib import Path
import re
from datetime import datetime, time
import pandas as pd
try:
    import pdfplumber
except ImportError:
    pdfplumber = None
import pypdf
from docx import Document

DAYS = {
    "MONDAY": "MON", "TUESDAY": "TUE", "WEDNESDAY": "WED", "THURSDAY": "THU",
    "FRIDAY": "FRI", "SATURDAY": "SAT", "SUNDAY": "SUN",
    "MON": "MON", "TUE": "TUE", "WED": "WED", "THU": "THU",
    "FRI": "FRI", "SAT": "SAT", "SUN": "SUN"
}

THEORY_TIMES = [
    ("09:00", "09:50"), ("09:55", "10:45"), ("10:50", "11:40"), ("11:45", "12:35"),
    ("13:15", "14:05"), ("14:10", "15:00"), ("15:05", "15:55"), ("16:00", "16:50"),
]
LAB_TIMES = [
    ("09:00", "10:40"), ("10:50", "12:30"), ("13:15", "14:55"), ("15:05", "16:45"),
]
PERIOD_TIMES = {
    1: (time(9, 0), time(9, 50)),
    2: (time(9, 55), time(10, 45)),
    3: (time(10, 50), time(11, 40)),
    4: (time(11, 45), time(12, 35)),
    5: (time(13, 15), time(14, 5)),
    6: (time(14, 10), time(15, 0)),
    7: (time(15, 5), time(15, 55)),
    8: (time(16, 0), time(16, 50)),
}
FREE_TOKENS = {"", "-", "free", "nil", "none", "nan", "available", "empty", "off", "no class"}
SLOT_TO_TIME = dict(zip(list("ABCDEFGH"), THEORY_TIMES))


def clean(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _read_csv_robust(path):
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        try:
            return pd.read_csv(path, encoding=enc, engine="python", on_bad_lines="skip", dtype=str)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        raise ValueError(f"Could not decode CSV file: {last_error}")
    return pd.read_csv(path, engine="python", on_bad_lines="skip", dtype=str)


def _read_legacy_word(path):
    import subprocess
    for cmd in [["antiword", str(path)], ["catdoc", str(path)]]:
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, encoding="cp1252", errors="replace", timeout=30)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except (FileNotFoundError, subprocess.SubprocessError):
            pass
    raise ValueError("This is an old .doc Word file. Please save it as .docx and upload again.")


def parse_time(value):
    s = clean(value).upper().replace(".", ":")
    s = s.replace("AM", " AM").replace("PM", " PM")
    s = re.sub(r"\s+", " ", s)
    for fmt in ["%H:%M", "%I:%M %p", "%I %p", "%H.%M"]:
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})\s*[:.]\s*(\d{2})\s*(AM|PM)?", s)
    if m:
        h, mi, ap = int(m.group(1)), int(m.group(2)), m.group(3)
        if ap == "PM" and h < 12:
            h += 12
        if ap == "AM" and h == 12:
            h = 0
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return time(h, mi)
    return None


def parse_date(value):
    s = clean(value)
    for fmt in [
        "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d %b %Y", "%d %B %Y"
    ]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()
        except ValueError:
            pass
    m_iso = re.search(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", s)
    if m_iso:
        try:
            return datetime(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))).date()
        except ValueError:
            pass
    return None


def _ocr_pdf(path):
    import fitz
    import pytesseract
    from PIL import Image
    import io
    output = []
    doc = fitz.open(path)
    try:
        for page_no, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(image, config="--psm 6")
            if text.strip():
                output.append(f"### OCR PAGE {page_no}\n{text.strip()}")
    finally:
        doc.close()
    return "\n\n".join(output)


def read_file(path):
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return _read_csv_robust(path).fillna("").astype(str).to_csv(index=False)
    if ext in [".xlsx", ".xls"]:
        output = []
        for name, df in pd.read_excel(path, sheet_name=None).items():
            output.append(f"### SHEET: {name}")
            output.append(df.fillna("").astype(str).to_csv(index=False))
        return "\n".join(output)
    if ext == ".docx":
        doc = Document(path)
        output = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                output.append(" | ".join(clean(c.text) for c in row.cells))
        return "\n".join(output)
    if ext == ".doc":
        return _read_legacy_word(path)
    if ext == ".pdf":
        output = []
        if pdfplumber is not None:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    if text.strip():
                        output.append(text)
                    for table in page.extract_tables() or []:
                        for row in table:
                            output.append(" | ".join(clean(c) for c in row))
        else:
            reader = pypdf.PdfReader(path)
            for page in reader.pages:
                text = page.extract_text() or ""
                if text.strip():
                    output.append(text)
        text = "\n".join(output).strip()
        if text:
            return text
        try:
            return _ocr_pdf(path)
        except Exception:
            return ""
    return Path(path).read_text(errors="ignore")


def preview(path):
    text = read_file(path)
    return {
        "filename": Path(path).name,
        "characters": len(text),
        "lines": len(text.splitlines()),
        "preview": text[:7000]
    }


def _extract_tables_as_grids(path):
    """Extract all 2D tables across Excel, CSV, Word, or PDF into uniform lists of lists."""
    ext = Path(path).suffix.lower()
    grids = []
    if ext == ".csv":
        df = _read_csv_robust(path).fillna("")
        if not df.empty:
            grids.append([df.columns.tolist()] + df.values.tolist())
    elif ext in [".xlsx", ".xls"]:
        for _, df in pd.read_excel(path, sheet_name=None).items():
            df = df.fillna("")
            if not df.empty:
                grids.append([df.columns.tolist()] + df.values.tolist())
    elif ext == ".docx":
        doc = Document(path)
        for t in doc.tables:
            grid = [[clean(c.text) for c in r.cells] for r in t.rows]
            if grid:
                grids.append(grid)
    elif ext == ".pdf":
        if pdfplumber is not None:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    for t in page.extract_tables() or []:
                        grid = [[clean(c) for c in r] for r in t]
                        if grid:
                            grids.append(grid)
    return grids


def _extract_course_code(text):
    m = re.search(r"\b([A-Z]{2,6}\d{3,5})\b", clean(text).upper())
    return m.group(1) if m else ""


def _slot_tokens(text):
    text = clean(text).upper()
    return re.findall(r"(?<![A-Z0-9])(L\d+(?:\s*\+\s*L\d+)*|[A-H](?:\s*\+\s*[A-H])?)(?![A-Z0-9])", text)


def _normalize_slot_token(token):
    return re.sub(r"\s+", "", token.upper())


def _parse_day_and_period_from_label(label):
    """Parse flexible header labels like 'Mon P1', 'Monday 2', 'Tue 09:00-10:00', 'Wednesday'."""
    s = clean(label)
    up = s.upper()
    day = None
    for k, v in DAYS.items():
        if re.search(rf"\b{k}\b", up) or k in up:
            day = v
            break
    # Check explicit time range inside label
    tm = re.search(r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)", s, re.I)
    if tm:
        st, et = parse_time(tm.group(1)), parse_time(tm.group(2))
        if st and et:
            return day, st, et

    # Check period number e.g. P1, Period 2, Hour 3, or standalone digit
    pm = re.search(r"(?:P|PERIOD|HOUR|HR|SLOT)?\s*(\d{1,2})\b", up)
    if pm:
        pno = int(pm.group(1))
        if pno in PERIOD_TIMES:
            st, et = PERIOD_TIMES[pno]
            return day, st, et

    # If only weekday is given (e.g. 'Monday' column covering full day), return day with None times
    return day, None, None


def parse_mca_faculty_docx(path):
    doc = Document(path)
    if len(doc.tables) < 2:
        raise ValueError("Not an MCA multi-table slot format.")

    grid = doc.tables[0]
    course_table = doc.tables[1]
    course_map = {}
    faculty_names = set()

    for row in course_table.rows[1:]:
        vals = [clean(c.text) for c in row.cells]
        if len(vals) < 8:
            continue
        slots, code, title, ctype, venue, faculty = vals[1], vals[2], vals[3], vals[5], vals[6], vals[7]
        if not faculty:
            continue
        faculty_names.add(faculty)
        metadata = {"code": code, "title": title, "type": ctype, "venue": venue, "faculty": faculty}
        for raw in re.split(r"[,;&/]", slots):
            for token in _slot_tokens(raw):
                course_map[_normalize_slot_token(token)] = metadata
        for token in _slot_tokens(slots):
            course_map[_normalize_slot_token(token)] = metadata

    if not faculty_names:
        raise ValueError("No faculty mappings found in second table.")

    day_rows = []
    for r in grid.rows[2:]:
        vals = [clean(c.text) for c in r.cells]
        if len(vals) < 3:
            continue
        day = DAYS.get(vals[0].upper(), "")
        if not day:
            continue
        day_rows.append((day, vals[1].upper(), vals))

    busy = []
    for day, kind, vals in day_rows:
        theory = kind == "THEORY"
        for i, cell in enumerate(vals[2:]):
            if i == 4:
                continue
            if theory:
                if i >= 8:
                    continue
                time_index = i if i < 4 else i - 1
                if time_index >= len(THEORY_TIMES):
                    continue
                st, et = THEORY_TIMES[time_index]
                parsed_st, parsed_et = parse_time(st), parse_time(et)
                tokens = _slot_tokens(cell)
                if not tokens and clean(cell).upper() in list("ABCDEFGH"):
                    tokens = [clean(cell).upper()]
                for token in tokens:
                    meta = course_map.get(_normalize_slot_token(token))
                    if meta:
                        busy.append({
                            "faculty": meta["faculty"], "weekday": day, "start": parsed_st, "end": parsed_et,
                            "source_slot": token, "course_code": meta["code"], "course_title": meta["title"]
                        })
                        break
            else:
                if i >= 8:
                    continue
                if i < 2:
                    block = 0
                elif i < 4:
                    block = 1
                elif 5 <= i < 7:
                    block = 2
                else:
                    block = 3
                st, et = LAB_TIMES[block]
                tokens = _slot_tokens(cell)
                for token in tokens:
                    meta = course_map.get(_normalize_slot_token(token))
                    if meta:
                        busy.append({
                            "faculty": meta["faculty"], "weekday": day, "start": parse_time(st), "end": parse_time(et),
                            "source_slot": token, "course_code": meta["code"], "course_title": meta["title"]
                        })
                        break

    seen = set()
    dedup = []
    for x in busy:
        key = (x["faculty"].lower(), x["weekday"], x["start"], x["end"], x["source_slot"], x["course_code"])
        if key not in seen:
            seen.add(key)
            dedup.append(x)
    return sorted(dedup, key=lambda x: (x["weekday"], x["start"], x["faculty"])), sorted(faculty_names)


def _extract_faculty_from_class_cell(cell_text):
    """
    Extract faculty initials/names from class-wise timetable cells like:
    - 'SE-BS' -> ['BS']
    - 'QT(RV)' -> ['RV']
    - 'DA(UMA)' -> ['UMA']
    - 'DA LAB - UMA & SN (8:45 AM - 10:45 AM)' -> ['UMA', 'SN']
    - 'WEB LAB - NV & PR (1:45 PM - 3:45 PM)' -> ['NV', 'PR']
    Ignores non-teaching entries like 'FREE', 'LIBRARY', 'PROJECT', 'PLACEMENT', 'LUNCH', 'BREAK'.
    """
    s = clean(cell_text)
    if not s:
        return []
    up = s.upper()
    ignore_exact = {"FREE", "LIBRARY", "PROJECT", "PLACEMENT", "LUNCH", "BREAK", "SHORT BREAK", "LUNCH BREAK", "-", "NIL", "NONE"}
    if up in ignore_exact or up.startswith("PROJECT ") or up.startswith("PLACEMENT "):
        return []

    # Strip trailing time ranges like '(8:45 AM - 10:45 AM)'
    s_no_time = re.sub(r"\(\s*\d{1,2}[:.]\d{2}.*?\)", "", s).strip()

    fac_part = ""
    # Pattern 1: 'SUBJECT - FAC1 & FAC2' or 'SE-BS'
    if "-" in s_no_time:
        parts = [p.strip() for p in s_no_time.split("-") if p.strip()]
        if len(parts) >= 2:
            fac_part = parts[-1]
    # Pattern 2: 'QT(RV)' or 'DA(UMA)'
    if not fac_part:
        m_paren = re.search(r"\(([A-Za-z.\s&,/]+)\)", s_no_time)
        if m_paren:
            fac_part = m_paren.group(1).strip()

    if not fac_part:
        return []

    # Split multiple faculty like 'UMA & SN' or 'NV & PR'
    raw_names = re.split(r"\s*(?:&|/|\+|,|and)\s*", fac_part, flags=re.I)
    cleaned = []
    for n in raw_names:
        n_clean = clean(n).strip("()")
        if n_clean and n_clean.upper() not in ignore_exact and not re.fullmatch(r"\d+.*", n_clean):
            cleaned.append(n_clean)
    return cleaned



def _extract_faculty_codes_from_master_cell(cell_text, legend_map):
    """
    Extracts faculty initials/codes from Format 3 master timetable cells such as:
    - 'PST(AKN)' -> ['AKN']
    - 'ELE ( R )' -> ['R']
    - 'PST LAB(L1,L2)  3AKN/3PBS' -> ['AKN', 'PBS']
    - 'PST LAB(L3,BL)  3BDP/2WAK+1NSM' -> ['BDP', 'WAK', 'NSM']
    - 'PROJECT B1/B2 (3NSM/1RS+2AKN)' -> ['NSM', 'RS', 'AKN']
    - 'CV(MAHESH)' -> ['MAHESH']
    - 'E/M LAB(E(B1+B2)R+M,M(B3+B4)) (BL,IBL)' -> ['R', 'M']
    Maps codes via legend_map (e.g. 'AKN' -> 'PROF.AMITHA K N (AKN)') when available.
    """
    s = clean(cell_text)
    if not s:
        return []
    up = s.upper()
    ignore_tokens = {
        "FREE", "VAP", "CS", "DS", "MATHS", "LAB", "BL", "IBL", "L1", "L2", "L3",
        "B1", "B2", "B3", "B4", "PG", "MCA", "PROJ", "PROJECT", "LAB1", "LAB2", "LAB3",
        "XXX", "E", "M", "K", "H", "S", "PST", "ELE", "OS", "PY", "ENG", "CN", "DBMS",
        "SE", "WP", "QT", "DA", "CA", "CV", "KAN", "SKILL", "INSIDE", "WITH", "THEN", "MOVE"
    }

    found_codes = []

    # 1. Extract codes inside parentheses like PST(AKN), ELE ( R ), CV(MAHESH)
    for grp in re.findall(r"\(([^()]+)\)", up):
        # Remove batch/room tokens like B1, B2, L1, L2, L3, BL, IBL, MCA LAB, PG LAB
        grp_clean = re.sub(r"\b(?:B[1-4]|L[1-3]|I?BL|MCA\s*LAB|PG\s*LAB|LAB\s*[1-3]?)\b", " ", grp)
        # Extract alpha tokens (with optional leading batch count digit like 3AKN, 2WAK, 1NSM)
        for tok in re.findall(r"(?:\d+)?([A-Z]{2,10}|R|M|D)", grp_clean):
            if tok not in ignore_tokens:
                found_codes.append(tok)

    # 2. Extract trailing batch-faculty patterns outside parentheses like '3AKN/3PBS' or '3BDP/2WAK+1NSM'
    for tok in re.findall(r"\b\d([A-Z]{2,6})\b", up):
        if tok not in ignore_tokens:
            found_codes.append(tok)

    # Deduplicate while preserving order and resolve full name from legend_map if present
    resolved = []
    seen = set()
    for code in found_codes:
        if code in seen:
            continue
        seen.add(code)
        full_name = legend_map.get(code, code)
        resolved.append(full_name)
    return resolved


def _parse_master_college_timetable_grid(grid, warnings):
    """
    Parses Format 3 (Master Multi-Class College Timetable like 'BCA & BSC ODD TT_ED1.xlsx'):
    - Header row contains 'Class' and time slot headers ('9:00 - 9:50', '9:50 - 10:40', '12:00 - 12:50', etc.)
    - Column 0 has Day headers ('Monday ...', 'Tuesday ...', ...) spanning multiple class rows below it
    - Bottom section may include a Faculty Legend mapping 'FULL NAME(CODE)'
    """
    if len(grid) < 5:
        return [], set()

    # Locate the header row containing 'Class' and time columns
    header_idx = None
    class_col = None
    time_cols = []  # list of (col_idx, start_time, end_time)

    for r_i, row in enumerate(grid[:15]):
        row_clean = [clean(c) for c in row]
        c_idx = None
        t_cols = []
        for idx, val in enumerate(row_clean):
            if val.lower() == "class":
                c_idx = idx
            tm = re.search(
                r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
                val, re.I
            )
            if tm and "break" not in val.lower() and "lunch" not in val.lower():
                st, et = parse_time(tm.group(1)), parse_time(tm.group(2))
                if st and et:
                    # Convert afternoon college hours (01:00 to 06:59 without explicit AM) to 24-hour PM (13:00 to 18:59)
                    if "am" not in tm.group(1).lower() and 1 <= st.hour <= 6:
                        st = st.replace(hour=st.hour + 12)
                    if "am" not in tm.group(2).lower() and 1 <= et.hour <= 6:
                        et = et.replace(hour=et.hour + 12)
                    t_cols.append((idx, st, et))
        if c_idx is not None and len(t_cols) >= 3:
            header_idx = r_i
            class_col = c_idx
            time_cols = t_cols
            break

    if header_idx is None:
        return [], set()

    # Build Faculty Legend Map from bottom rows (e.g., 'PROF.AMITHA K N(AKN)' -> {'AKN': 'PROF.AMITHA K N (AKN)'})
    legend_map = {}
    legend_start_row = len(grid)
    for r_i in range(header_idx + 1, len(grid)):
        row_clean = [clean(c) for c in grid[r_i]]
        if any("class teachers" in c.lower() or "faculty member" in c.lower() for c in row_clean):
            legend_start_row = r_i
        for cell in row_clean:
            m_leg = re.search(r"([A-Za-z.\s]+?)\s*\(\s*([A-Z]{2,6})\s*\)", cell)
            if m_leg and any(k in cell.upper() for k in ["PROF", "DR", "MR", "MS", "MRS", "SIR", "MAM"]) or (
                m_leg and r_i >= legend_start_row
            ):
                fname = clean(m_leg.group(1))
                fcode = m_leg.group(2).strip().upper()
                if len(fname) >= 3:
                    legend_map[fcode] = f"{fname} ({fcode})"

    records = []
    faculty_names = set(legend_map.values())
    current_day = None

    # Sort time_cols by column index
    time_cols.sort(key=lambda x: x[0])

    for r_i in range(header_idx + 1, legend_start_row):
        row = grid[r_i]
        if not row:
            continue
        first_cell = clean(row[0]).upper()
        for d_key, d_code in DAYS.items():
            if first_cell.startswith(d_key):
                current_day = d_code
                break

        if not current_day:
            continue

        class_val = clean(row[class_col]) if class_col < len(row) else ""
        if not class_val or class_val.lower() == "class":
            continue

        for pos, (c_idx, st, et) in enumerate(time_cols):
            cell_val = clean(row[c_idx]) if c_idx < len(row) else ""
            if not cell_val:
                continue

            # If this is a LAB / PROJECT spanning multiple empty adjacent period cells in the same session, extend end_time
            eff_et = et
            if any(k in cell_val.upper() for k in ["LAB", "PROJECT", "PROJ"]):
                next_pos = pos + 1
                while next_pos < len(time_cols):
                    n_cidx, _, n_et = time_cols[next_pos]
                    # Stop if crossing a room column gap (more than 1 column jump indicates Afternoon session block)
                    if n_cidx - time_cols[next_pos - 1][0] > 2:
                        break
                    n_val = clean(row[n_cidx]) if n_cidx < len(row) else ""
                    if not n_val:
                        eff_et = n_et
                        next_pos += 1
                    else:
                        break

            facs = _extract_faculty_codes_from_master_cell(cell_val, legend_map)
            for fac in facs:
                faculty_names.add(fac)
                records.append({
                    "faculty": fac,
                    "weekday": current_day,
                    "start": st,
                    "end": eff_et,
                    "source_slot": class_val,
                    "course_code": _extract_course_code(cell_val),
                    "course_title": f"{class_val}: {cell_val}",
                })

    # Deduplicate identical faculty + weekday + start + end
    seen = set()
    dedup = []
    for r in records:
        key = (r["faculty"], r["weekday"], r["start"], r["end"])
        if key not in seen:
            seen.add(key)
            dedup.append(r)

    return dedup, faculty_names


def _parse_class_wise_timetable_grid(grid, warnings):
    """
    Handles Class-Wise Weekly Timetables where:
    - First column is Weekday (MONDAY, TUESDAY, ...)
    - Top header row is Time Ranges ('8:45 AM - 9:45 AM', '9:45 AM - 10:45 AM', ...)
    - Cells contain 'SUBJECT-FACULTY' or 'SUBJECT(FACULTY)'
    """
    if len(grid) < 2 or len(grid[0]) < 2:
        return [], set()

    header_row = [clean(c) for c in grid[0]]
    first_col = [clean(r[0]) for r in grid if r]

    first_col_days = sum(1 for c in first_col[1:] if DAYS.get(c.upper()) or DAYS.get(c.upper()[:3]))
    # Check if top header has time ranges (like '8:45 AM - 9:45 AM')
    header_time_cols = []
    for idx, h in enumerate(header_row[1:], start=1):
        tm = re.search(
            r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
            h, re.I
        )
        if tm:
            st, et = parse_time(tm.group(1)), parse_time(tm.group(2))
            if st and et:
                header_time_cols.append((idx, st, et))

    if first_col_days < 2 or len(header_time_cols) < 2:
        return [], set()

    records = []
    faculty_names = set()

    for row in grid[1:]:
        if not row:
            continue
        raw_day = clean(row[0]).upper()
        day = DAYS.get(raw_day) or DAYS.get(raw_day[:3])
        if not day:
            continue

        for col_idx, st, et in header_time_cols:
            cell_val = clean(row[col_idx]) if col_idx < len(row) else ""
            facs = _extract_faculty_from_class_cell(cell_val)
            for fac in facs:
                faculty_names.add(fac)
                records.append({
                    "faculty": fac,
                    "weekday": day,
                    "start": st,
                    "end": et,
                    "source_slot": "",
                    "course_code": _extract_course_code(cell_val),
                    "course_title": cell_val,
                })

    return records, faculty_names


def _parse_matrix_faculty_grid(grid, warnings):
    """
    Universal parser for ANY college timetable matrix:
    - College A: Faculty names in rows, Weekdays/Periods in columns
    - College B: Faculty names in columns, Weekdays/Periods in rows (auto-transposed)
    """
    if len(grid) < 2 or len(grid[0]) < 2:
        return [], set()

    header_row = [clean(c) for c in grid[0]]
    first_col = [clean(r[0]) for r in grid if r]

    header_has_days = any(_parse_day_and_period_from_label(c)[0] is not None for c in header_row[1:])
    col_has_days = any(_parse_day_and_period_from_label(c)[0] is not None for c in first_col[1:])

    # College B format: Days/Periods are in the first column and Faculty names are in the top header
    if col_has_days and not header_has_days:
        grid = [list(col) for col in zip(*grid)]
        header_row = [clean(c) for c in grid[0]]

    if not any(_parse_day_and_period_from_label(c)[0] is not None for c in header_row[1:]):
        return [], set()

    records = []
    faculty_names = set()
    col_mappings = []
    for col_idx, h in enumerate(header_row[1:], start=1):
        day, st, et = _parse_day_and_period_from_label(h)
        if not day:
            warnings.append(f"Column '{h}': Could not confidently identify weekday/period; skipped.")
            col_mappings.append(None)
        else:
            # Default to full-day period 1-8 if only weekday name was given (e.g. 'Monday')
            if st is None or et is None:
                st, et = time(9, 0), time(17, 0)
            col_mappings.append((col_idx, day, st, et))

    for row_idx, row in enumerate(grid[1:], start=2):
        fac = clean(row[0]) if row else ""
        if not fac or fac.lower() in {"faculty", "faculty name", "name", "staff", "lecturer", "nan", "none"}:
            warnings.append(f"Row {row_idx}: Missing or invalid faculty name; skipped.")
            continue
        faculty_names.add(fac)
        for mapping in col_mappings:
            if not mapping:
                continue
            col_idx, day, st, et = mapping
            cell_val = clean(row[col_idx]) if col_idx < len(row) else ""
            if cell_val.lower() not in FREE_TOKENS:
                records.append({
                    "faculty": fac,
                    "weekday": day,
                    "start": st,
                    "end": et,
                    "source_slot": "",
                    "course_code": _extract_course_code(cell_val),
                    "course_title": cell_val,
                })
    return records, faculty_names


def parse_faculty_file(path, warnings=None):
    if warnings is None:
        warnings = []
    ext = Path(path).suffix.lower()

    # 1. Try specialized multi-table DOCX format first (e.g., MCA 2025.docx)
    if ext == ".docx":
        try:
            return parse_mca_faculty_docx(path)
        except ValueError:
            pass

    records = []
    faculty_names = set()
    grids = _extract_tables_as_grids(path)

    # 2. Try row-based tabular format (Faculty, Day/Date, Start, End, Status)
    for grid in grids:
        if len(grid) < 2:
            continue
        headers = [clean(c).lower() for c in grid[0]]
        col_idx = {h: i for i, h in enumerate(headers)}

        def find_col(*names):
            for n in names:
                if n in col_idx:
                    return col_idx[n]
            return None

        fcol = find_col("faculty", "faculty name", "name", "lecturer", "staff name")
        dcol = find_col("day", "weekday", "date")
        scol = find_col("start", "start time", "from")
        ecol = find_col("end", "end time", "to")
        statcol = find_col("status", "availability", "state")
        ccol = find_col("course code", "code")
        ctcol = find_col("course", "course title", "subject")

        if fcol is not None and dcol is not None and scol is not None and ecol is not None:
            for r_i, row in enumerate(grid[1:], start=2):
                fac = clean(row[fcol]) if fcol < len(row) else ""
                if not fac:
                    warnings.append(f"Row {r_i}: Missing faculty name; skipped.")
                    continue
                raw_day = clean(row[dcol]) if dcol < len(row) else ""
                day = DAYS.get(raw_day.upper()) or DAYS.get(raw_day.upper()[:3])
                if not day:
                    dt = parse_date(raw_day)
                    if dt:
                        day = dt.strftime("%a").upper()[:3]
                if not day:
                    warnings.append(f"Row {r_i} ({fac}): Could not identify weekday/date '{raw_day}'.")
                    continue
                st = parse_time(row[scol]) if scol < len(row) else None
                et = parse_time(row[ecol]) if ecol < len(row) else None
                if not st or not et:
                    warnings.append(f"Row {r_i} ({fac}): Invalid start/end timing.")
                    continue

                faculty_names.add(fac)
                status_val = clean(row[statcol]).upper() if (statcol is not None and statcol < len(row)) else "BUSY"
                if status_val in {"FREE", "AVAILABLE", "EMPTY", "OFF"}:
                    continue  # Faculty is free; do not add a BUSY slot

                records.append({
                    "faculty": fac,
                    "weekday": day,
                    "start": st,
                    "end": et,
                    "source_slot": "",
                    "course_code": clean(row[ccol]) if (ccol is not None and ccol < len(row)) else "",
                    "course_title": clean(row[ctcol]) if (ctcol is not None and ctcol < len(row)) else "",
                })

    # 3. Try Format 3: Master Multi-Class College Timetable (e.g. BCA & BSC ODD TT with merged Day blocks & Faculty Legend)
    if not faculty_names:
        for grid in grids:
            m3_recs, m3_facs = _parse_master_college_timetable_grid(grid, warnings)
            if m3_facs:
                records.extend(m3_recs)
                faculty_names.update(m3_facs)

    # 4. Try Format 2: Class-Wise Weekly Timetable format (Days in rows, Time slots in headers, cells like 'SE-BS' or 'DA LAB - UMA & SN')
    if not faculty_names:
        for grid in grids:
            c_recs, c_facs = _parse_class_wise_timetable_grid(grid, warnings)
            if c_facs:
                records.extend(c_recs)
                faculty_names.update(c_facs)

    # 4. Try Matrix format (College A: Faculty in rows, Days/Periods in cols OR College B: Faculty in cols, Periods in rows)
    if not faculty_names:
        for grid in grids:
            m_recs, m_facs = _parse_matrix_faculty_grid(grid, warnings)
            if m_facs:
                records.extend(m_recs)
                faculty_names.update(m_facs)

    if not faculty_names:
        raise ValueError("Could not identify faculty names, weekdays, or periods in the uploaded faculty timetable.")

    return records, sorted(faculty_names)


def _exam_records_from_text(text):
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    exams = []
    current_date = None
    for line in lines:
        dates = re.findall(r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})\b", line)
        if dates:
            for ds in dates:
                current_date = parse_date(ds)
                if current_date:
                    break

        # Check explicit time range on the same line as date
        t_match = re.search(
            r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
            line, re.I
        )
        if current_date and t_match:
            st, et = parse_time(t_match.group(1)), parse_time(t_match.group(2))
            if st and et:
                slot_m = re.search(r"\b(?:Slot\s*)?([A-H]|FN|AN)\b", line, re.I)
                slot_lbl = slot_m.group(1).upper() if slot_m else ("FN" if st.hour < 12 else "AN")
                exams.append({
                    "date": current_date, "start": st, "end": et,
                    "slot": slot_lbl, "course": line[:80],
                    "course_code": _extract_course_code(line), "class_name": ""
                })
                continue

        slots = re.findall(r"(?<![A-Z0-9])([A-H])(?![A-Z0-9])", line.upper())
        if current_date and len(slots) >= 2:
            picked = []
            for x in slots:
                if x not in picked:
                    picked.append(x)
            for idx, slot in enumerate(picked[:2]):
                st, et = (time(10, 0), time(12, 0)) if idx == 0 else (time(14, 0), time(16, 0))
                exams.append({
                    "date": current_date, "start": st, "end": et,
                    "slot": slot, "course": "Examination",
                    "course_code": "", "class_name": ""
                })
            current_date = None

    # Handle PDF text where Date, Start Time, End Time appear on consecutive lines
    if not exams:
        idx = 0
        while idx < len(lines):
            d = parse_date(lines[idx])
            if d and idx + 2 < len(lines):
                st = parse_time(lines[idx + 1])
                et = parse_time(lines[idx + 2])
                if st and et:
                    course_lbl = lines[idx + 3] if (idx + 3 < len(lines) and not parse_date(lines[idx + 3])) else "Examination"
                    exams.append({
                        "date": d, "start": st, "end": et,
                        "slot": "FN" if st.hour < 12 else "AN",
                        "course": course_lbl[:80], "course_code": _extract_course_code(course_lbl), "class_name": ""
                    })
                    idx += 3
                    continue
            idx += 1

    if not exams:
        all_dates = []
        for m in re.finditer(r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})\b", text):
            d = parse_date(m.group(0))
            if d:
                all_dates.append(d)
        slots = [x for x in re.findall(r"(?<![A-Z0-9])([A-H])(?![A-Z0-9])", text.upper()) if x in "ABCDEFGH"]
        if all_dates:
            for i, d in enumerate(all_dates):
                pair = slots[i * 2:i * 2 + 2]
                if len(pair) >= 2:
                    for j, slot in enumerate(pair):
                        st, et = (time(10, 0), time(12, 0)) if j == 0 else (time(14, 0), time(16, 0))
                        exams.append({
                            "date": d, "start": st, "end": et, "slot": slot,
                            "course": "Examination", "course_code": "", "class_name": ""
                        })

    seen = set()
    out = []
    for e in exams:
        k = (e["date"], e["start"], e["end"], e["slot"])
        if k not in seen:
            seen.add(k)
            out.append(e)
    return out


def _parse_exam_grid(grid, warnings):
    """Parse exam tables across any column arrangement (Date, Session/Time, Slot, Course, etc.)."""
    if len(grid) < 2:
        return []
    headers = [clean(c).lower() for c in grid[0]]
    col_idx = {h: i for i, h in enumerate(headers)}

    def find_col(*names):
        for n in names:
            for h, idx in col_idx.items():
                if n == h or n in h:
                    return idx
        return None

    dcol = find_col("date", "exam date", "day/date")
    scol = find_col("start", "from")
    ecol = find_col("end", "to")
    sesscol = find_col("session", "timing", "time")
    slotcol = find_col("slot")
    coursecol = find_col("course name", "course", "subject")
    codecol = find_col("course code", "code")
    classcol = find_col("class", "batch", "program")

    slot_times = {
        "A": (time(10, 0), time(12, 0)), "B": (time(14, 0), time(16, 0)),
        "C": (time(10, 0), time(12, 0)), "D": (time(14, 0), time(16, 0)),
        "E": (time(10, 0), time(12, 0)), "F": (time(14, 0), time(16, 0)),
        "G": (time(10, 0), time(12, 0)), "H": (time(14, 0), time(16, 0)),
        "FN": (time(9, 30), time(12, 30)), "AN": (time(13, 30), time(16, 30)),
    }

    # Check if headers themselves contain Forenoon / Afternoon columns (like PG_SS_2025-26AY_CA-1_TT.docx)
    fn_col = find_col("forenoon", "fn", "morning")
    an_col = find_col("afternoon", "an", "evening")

    out = []
    for r_i, row in enumerate(grid[1:], start=2):
        cells = [clean(c) for c in row]
        if not any(cells):
            continue

        d = parse_date(cells[dcol]) if (dcol is not None and dcol < len(cells)) else None
        if not d:
            for c in cells:
                d = parse_date(c)
                if d:
                    break
        if not d:
            warnings.append(f"Exam Row {r_i}: Could not confidently identify exam date in '{' | '.join(cells)[:60]}'.")
            continue

        if fn_col is not None or an_col is not None:
            for col_p, default_sess, default_times in [
                (fn_col, "FN", (time(10, 0), time(12, 0))),
                (an_col, "AN", (time(14, 0), time(16, 0)))
            ]:
                if col_p is not None and col_p < len(cells) and cells[col_p] and cells[col_p] != "-":
                    val = cells[col_p]
                    sm = re.search(r"Slot\s*([A-H])", val, re.I)
                    slot_lbl = sm.group(1).upper() if sm else default_sess
                    out.append({
                        "date": d, "start": default_times[0], "end": default_times[1],
                        "slot": slot_lbl, "course": val,
                        "course_code": _extract_course_code(val), "class_name": ""
                    })
            continue

        st = parse_time(cells[scol]) if (scol is not None and scol < len(cells)) else None
        et = parse_time(cells[ecol]) if (ecol is not None and ecol < len(cells)) else None

        if (not st or not et) and sesscol is not None and sesscol < len(cells):
            tm = re.search(
                r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
                cells[sesscol], re.I
            )
            if tm:
                st, et = parse_time(tm.group(1)), parse_time(tm.group(2))

        slot_raw = clean(cells[slotcol]).upper() if (slotcol is not None and slotcol < len(cells)) else ""
        slot_m = re.search(r"(?:SLOT\s*)?([A-H]|FN|AN)", slot_raw)
        slot_lbl = slot_m.group(1) if slot_m else slot_raw[:10]

        if (not st or not et) and slot_lbl in slot_times:
            st, et = slot_times[slot_lbl]

        if not st or not et:
            # Scan entire row for a time range before warning
            joined = " ".join(cells)
            tm = re.search(
                r"(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)\s*(?:-|to|–)\s*(\d{1,2}[:.]\d{2}\s*(?:AM|PM)?)",
                joined, re.I
            )
            if tm:
                st, et = parse_time(tm.group(1)), parse_time(tm.group(2))

        if not st or not et:
            warnings.append(f"Exam Row {r_i} ({d}): Could not identify session timing.")
            continue

        out.append({
            "date": d,
            "start": st,
            "end": et,
            "slot": slot_lbl or ("FN" if st.hour < 12 else "AN"),
            "course": clean(cells[coursecol]) if (coursecol is not None and coursecol < len(cells)) else "Exam",
            "course_code": clean(cells[codecol]) if (codecol is not None and codecol < len(cells)) else "",
            "class_name": clean(cells[classcol]) if (classcol is not None and classcol < len(cells)) else "",
        })
    return out


def parse_exam_file(path, warnings=None):
    if warnings is None:
        warnings = []
    grids = _extract_tables_as_grids(path)
    exams = []
    for grid in grids:
        exams.extend(_parse_exam_grid(grid, warnings))

    if not exams:
        text = read_file(path)
        exams = _exam_records_from_text(text)

    # Deduplicate exact same date+start+end+slot
    seen = set()
    dedup = []
    for e in exams:
        key = (e["date"], e["start"], e["end"], e["slot"])
        if key not in seen:
            seen.add(key)
            dedup.append(e)

    if not dedup:
        raise ValueError("Could not identify exam dates, sessions, or timings in the uploaded exam file.")
    return dedup


def parse_room_file(path, warnings=None):
    if warnings is None:
        warnings = []
    grids = _extract_tables_as_grids(path)
    rooms = []

    for grid in grids:
        if not grid:
            continue
        # Scan top 10 rows for Master Timetable room columns ('Morning session class room no', etc.)
        master_handled = False
        for hdr_r in range(min(10, len(grid))):
            row_h = [clean(c).lower() for c in grid[hdr_r]]
            m_rcols = [idx for idx, h in enumerate(row_h) if "room" in h]
            if m_rcols and any(h == "class" for h in row_h):
                for row in grid[hdr_r + 1:]:
                    row_str = " ".join(clean(x).lower() for x in row)
                    if "class teachers" in row_str or "faculty member" in row_str:
                        break
                    for mc in m_rcols:
                        if mc < len(row):
                            rv = clean(row[mc])
                            if rv and rv.lower() not in {"nan", "none", "room"}:
                                rooms.append(rv)
                master_handled = True
                break
        if master_handled:
            continue

        headers = [clean(c).lower() for c in grid[0]]
        rcol = None
        for idx, h in enumerate(headers):
            if any(k in h for k in ["room", "hall", "venue"]):
                rcol = idx
                break

        has_header = rcol is not None or any(k in " ".join(headers) for k in ["room", "capacity", "date", "building", "hall", "venue"])
        start_row = 1 if has_header else 0

        for r_i, row in enumerate(grid[start_row:], start=start_row + 1):
            cells = [clean(c) for c in row]
            if not any(cells):
                continue
            candidate = ""
            if rcol is not None and rcol < len(cells):
                c = cells[rcol]
                if c and c.lower() not in {"room", "room no", "room number", "nan", "none", "venue"}:
                    candidate = re.sub(r"^room\s*(?:no\.?|number)?\s*", "", c, flags=re.I).strip()
            elif len(cells) == 1:
                c = cells[0]
                if re.fullmatch(r"(?:Room\s*)?[A-Za-z]{0,3}-?\d{2,4}[A-Za-z]?", c, re.I):
                    candidate = re.sub(r"^room\s*", "", c, flags=re.I).strip()
            else:
                for c in cells:
                    if re.fullmatch(r"(?:Room\s*)?[A-Za-z]{0,3}-?\d{2,4}[A-Za-z]?", c, re.I) and not re.fullmatch(r"202\d", c):
                        candidate = re.sub(r"^room\s*", "", c, flags=re.I).strip()
                        break
            if candidate:
                rooms.append(candidate)
            elif rcol is not None:
                warnings.append(f"Room Row {r_i}: Could not confidently identify room number in '{' | '.join(cells)[:50]}'.")

    if not rooms:
        text = read_file(path)
        for line in text.splitlines():
            m = re.search(r"\b(?:Room\s*(?:No\.?)?\s*)?([A-Za-z]{0,3}-?\d{2,4}[A-Za-z]?)\b", line, re.I)
            if m and not re.search(r"202\d", m.group(1)):
                rooms.append(m.group(1))

    rooms = list(dict.fromkeys(rooms))
    if not rooms:
        raise ValueError("No room numbers could be identified in the room details file.")
    return rooms
