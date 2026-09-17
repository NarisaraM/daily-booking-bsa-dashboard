"""
Daily Booking Status vs BSA (Block Space Agreement) -- deterministic pipeline.

No AI is used anywhere in this script: extraction is done with the library
that matches each file's real format, and the BSA comparison is pure
arithmetic (sum + divide + threshold). This mirrors how the sibling
"Daily booking" project was built and keeps the numbers 100% reproducible.

All source files live in input/ (BSA.xlsx, the SKED schedule export, the
master booking export, and any future PDFs) -- drop a new day's exports in
there (removing or leaving the previous ones; the newest by mtime wins
either way) and re-run. Only the generated report/dashboard live alongside
this script itself.

Extraction (each file read with the correct library for its real format):
  - input/BSA.xlsx           openpyxl   -> BSA per-lane / per-port TEU allocation
  - input/*SKED*.xls          xlrd       -> vessel schedule (Service, vessel
                                            code/name, ETD, slot-share flag)
  - the master booking export xlrd       -> one row per booking line
                                            (VSL/VOY/POD/TEU/weight/SVC/...),
                                            auto-detected as the newest
                                            .xls/.xlsx in input/ that isn't
                                            BSA.xlsx or the SKED file
  - any input/*.pdf                       pdfplumber -> raw text only.
    pdfplumber can only open PDFs (it cannot read .xls/.xlsx binaries), so
    it is used exclusively for that role. No PDFs ship with this pipeline
    today, so this step is a no-op unless someone later drops a booking
    confirmation PDF here -- if they do, its text is dumped to
    pdf_text_extracts/<name>.txt instead of being silently ignored.

Analysis (pure arithmetic + lookups, no AI):
  - Each booking's POD is mapped to a BSA port group (VNSGN, HKHKG, CNXMN,
    CNSHK, TWKEL, KR, JP, RU, IDJKT, CNSHA -- everything else is tracked
    separately as OTHER/unmapped so it stays visible instead of vanishing).
  - Each booking's SVC is mapped to a BSA lane (identical to SVC except
    PCI2 -> PCI, per Report.xlsx's "Master BSA" mapping sheet).
  - Bookings are grouped by ISO week (Monday-Sunday) of ETD, then by lane:
    the BSA figures in BSA.xlsx are a per-lane WEEKLY quota, so all
    sailings of the same lane in the same week share one quota.
  - Booked TEU = sum of the master file's own TEU column (already verified
    to implement 20'=1 TEU, 40'/45'=2 TEU for every row in this dataset --
    see verify_teu_formula() below -- so it is used as-is rather than
    recomputed).
  - Status per number: OK (<100%), FULL (==100%), OVER (>100%).
  - BSA.xlsx carries no weight ceiling, so Total Booked Weight is reported
    as an informational figure only (no OK/OVER/FULL judgement).

Output (in this same folder):
  - Daily_Booking_Status_Report.xlsx  -- colour-coded Excel report
  - Daily_Booking_Dashboard.html      -- self-contained interactive
                                          dashboard, filterable by
                                          destination port, opens in any
                                          browser (double-click, no server)

Run:
    python booking_dashboard.py
"""
import glob
import json
import os
import re
from collections import defaultdict
from datetime import datetime, timedelta

import openpyxl
import xlrd

FOLDER = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(FOLDER, "input")
EXCEL_OUTPUT = os.path.join(FOLDER, "Daily_Booking_Status_Report.xlsx")
HTML_OUTPUT = os.path.join(FOLDER, "Daily_Booking_Dashboard.html")
# Also published as index.html so GitHub Pages (served from the repo root)
# always reflects the latest run once pushed.
INDEX_OUTPUT = os.path.join(FOLDER, "index.html")
PDF_TEXT_DIR = os.path.join(FOLDER, "pdf_text_extracts")

# BSA.xlsx port columns, in the order they appear in the sheet.
PORT_GROUPS = ["VNSGN", "HKHKG", "CNXMN", "CNSHK", "TWKEL", "KR", "JP", "RU", "IDJKT", "CNSHA"]
DIRECT_POD_GROUPS = {"VNSGN", "HKHKG", "CNXMN", "CNSHK", "TWKEL", "CNSHA", "IDJKT"}

# Raw SVC code (as it appears in the booking/schedule files) -> BSA lane
# label (as it appears in BSA.xlsx). Identity for every lane except this one
# (PCI2 is the raw service code; "PCI" is the BSA lane it draws from -- see
# Report.xlsx's "Master BSA" sheet, which documents this exact mapping).
SVC_TO_LANE = {"PCI2": "PCI"}

WEIGHT_DIVISOR = 1000.0  # booking weight is in kg; report weight in metric tons

# A vessel flagged in the SKED file's USED column (e.g. "BUY(PCS)") only
# bought a slot-share of the normal BSA from another carrier, so its usable
# TEU capacity -- both the total and each individual port's allowance -- is
# capped at this fraction of the full BSA. Weight is not affected.
SLOT_SHARE_TEU_FACTOR = 0.6

# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------

def _latest(paths):
    return max(paths, key=os.path.getmtime)


def find_bsa_file():
    candidates = glob.glob(os.path.join(INPUT_FOLDER, "*BSA*.xls*"))
    candidates = [p for p in candidates if not os.path.basename(p).startswith("~$")]
    if not candidates:
        raise FileNotFoundError(f"No BSA allocation file (name containing 'BSA') found in {INPUT_FOLDER}")
    return _latest(candidates)


def find_sked_file():
    candidates = [
        p for p in glob.glob(os.path.join(INPUT_FOLDER, "*.xls*"))
        if not os.path.basename(p).startswith("~$") and "sked" in os.path.basename(p).lower()
    ]
    if not candidates:
        raise FileNotFoundError(f"No schedule file (name containing 'SKED') found in {INPUT_FOLDER}")
    return _latest(candidates)


def find_master_booking_file():
    bsa_file = find_bsa_file()
    sked_file = find_sked_file()
    excluded_names = {"report.xlsx", os.path.basename(EXCEL_OUTPUT).lower()}
    candidates = [
        p for p in glob.glob(os.path.join(INPUT_FOLDER, "*.xls*"))
        if not os.path.basename(p).startswith("~$")
        and p != bsa_file
        and p != sked_file
        and os.path.basename(p).lower() not in excluded_names
    ]
    if not candidates:
        raise FileNotFoundError(f"No master booking export found in {INPUT_FOLDER}")
    return _latest(candidates)


def find_pdf_files():
    return [p for p in glob.glob(os.path.join(INPUT_FOLDER, "*.pdf"))]


# ---------------------------------------------------------------------------
# Extraction -- BSA.xlsx (openpyxl: real .xlsx file)
# ---------------------------------------------------------------------------

def parse_bsa(path):
    """Return dict lane -> {"total_teu": float, "ports": {port_group: teu}}."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    port_header = rows[1]
    port_cols = {}
    for c in range(2, len(port_header) - 1):
        name = port_header[c]
        if name:
            port_cols[c] = str(name).strip().upper()

    bsa = {}
    for row in rows[2:]:
        lane = row[0]
        if not lane:
            continue
        lane = str(lane).strip().upper()
        total_teu = float(row[1]) if row[1] else 0.0
        ports = {}
        for c, name in port_cols.items():
            val = row[c] if c < len(row) else None
            if val:
                ports[name] = float(val)
        bsa[lane] = {"total_teu": total_teu, "ports": ports}
    return bsa


# ---------------------------------------------------------------------------
# Extraction -- schedule file (xlrd: legacy .xls)
# ---------------------------------------------------------------------------

def _parse_xls_datetime(raw, workbook):
    if raw is None or raw == "":
        return None
    if isinstance(raw, (int, float)):
        try:
            return xlrd.xldate_as_datetime(raw, workbook.datemode)
        except Exception:
            return None
    text = str(raw).strip()
    for value, fmt in ((text, "%Y-%m-%d %H:%M"), (text[:10], "%Y-%m-%d")):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def parse_sked(path):
    """Return dict (vessel_code, voyage) -> {service, vessel_name, etd, slot_share, slot_share_label}.

    "Vyg Bound" (e.g. "2607N") is used as the voyage key -- it already
    combines the "Vyg" and "Bound" columns into the same format the master
    booking file's VOY column uses. A vessel/voyage appears once per
    port-of-call; the first row wins, since Service/Name/ETD/USED don't
    vary between a sailing's own port-of-call rows.
    """
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    header = [sh.cell_value(0, c) for c in range(sh.ncols)]
    idx = {h: c for c, h in enumerate(header)}

    lookup = {}
    for r in range(1, sh.nrows):
        code = sh.cell_value(r, idx["Vessel"])
        voy = sh.cell_value(r, idx["Vyg Bound"])
        if not code or not voy:
            continue
        key = (str(code).strip().upper(), str(voy).strip().upper())
        if key in lookup:
            continue

        used = sh.cell_value(r, idx["USED"])
        lookup[key] = {
            "service": str(sh.cell_value(r, idx["Service"])).strip().upper(),
            "vessel_name": str(sh.cell_value(r, idx["Vessel Name"])).strip(),
            "etd": _parse_xls_datetime(sh.cell_value(r, idx["ETD Date"]), wb),
            "slot_share": bool(str(used).strip()),
            "slot_share_label": str(used).strip(),
        }
    return lookup


# ---------------------------------------------------------------------------
# Extraction -- master booking export (xlrd: legacy .xls)
# ---------------------------------------------------------------------------

def verify_teu_formula(sh, idx):
    """Sanity-check that TEU already implements 20'=1 / 40'-45'=2 for this
    file, since that's the assumption build_analysis() relies on instead of
    recomputing TEU from the container-count columns itself."""
    checked = mismatches = 0
    for r in range(1, sh.nrows):
        if not sh.cell_value(r, idx["VSL"]):
            continue
        c20 = sh.cell_value(r, idx["C20"]) or 0
        c40 = sh.cell_value(r, idx["C40"]) or 0
        c45 = sh.cell_value(r, idx["C45"]) or 0
        teu = sh.cell_value(r, idx["TEU"]) or 0
        checked += 1
        if abs((c20 * 1 + c40 * 2 + c45 * 2) - teu) > 0.01:
            mismatches += 1
    return checked, mismatches


def parse_reefer_qty(raw):
    """RFCNT holds free-text like "45REx2" (container size x quantity),
    occasionally with more than one spec in the same cell. Sum every
    "x<number>" quantity found; blank/non-matching cells count as 0."""
    if not raw:
        return 0
    return sum(int(n) for n in re.findall(r"x\s*(\d+)", str(raw), flags=re.IGNORECASE))


def parse_bookings(path):
    """Return one dict per booking line from the master booking export."""
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    header = [sh.cell_value(0, c) for c in range(sh.ncols)]
    idx = {h: c for c, h in enumerate(header)}

    checked, mismatches = verify_teu_formula(sh, idx)
    if mismatches:
        print(f"  WARNING: TEU column does not match 20'=1/40'-45'=2 for {mismatches}/{checked} rows "
              f"-- booked TEU totals may be off. Recompute manually if this grows.")

    rows = []
    for r in range(1, sh.nrows):
        vsl = sh.cell_value(r, idx["VSL"])
        voy = sh.cell_value(r, idx["VOY"])
        if not vsl or not voy:
            continue
        etd_raw = sh.cell_value(r, idx["ETD"])
        rows.append({
            "vsl": str(vsl).strip().upper(),
            "voy": str(voy).strip().upper(),
            "svc": str(sh.cell_value(r, idx["SVC"]) or "").strip().upper(),
            "pod": str(sh.cell_value(r, idx["POD"]) or "").strip().upper(),
            "teu": float(sh.cell_value(r, idx["TEU"]) or 0),
            "weight_kg": float(sh.cell_value(r, idx["BK Tot Weight"]) or 0),
            "reefer_qty": parse_reefer_qty(sh.cell_value(r, idx["RFCNT"])),
            "etd": _parse_xls_datetime(etd_raw, wb),
        })
    return rows


# ---------------------------------------------------------------------------
# Extraction -- any PDFs dropped in input/ (pdfplumber: real PDFs only)
# ---------------------------------------------------------------------------

def extract_pdf_texts():
    pdf_paths = find_pdf_files()
    if not pdf_paths:
        print("No PDF files found in input/ -- pdfplumber extraction skipped.")
        return {}

    import pdfplumber  # imported lazily so the rest of the script works without it installed

    texts = {}
    os.makedirs(PDF_TEXT_DIR, exist_ok=True)
    for path in pdf_paths:
        name = os.path.basename(path)
        with pdfplumber.open(path) as pdf:
            text = "\n\n".join(page.extract_text() or "" for page in pdf.pages)
        texts[name] = text
        out_path = os.path.join(PDF_TEXT_DIR, os.path.splitext(name)[0] + ".txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"  Extracted text from {name} -> {out_path} ({len(text)} chars)")
    return texts


# ---------------------------------------------------------------------------
# Mapping helpers
# ---------------------------------------------------------------------------

def pod_group(pod):
    pod = (pod or "").strip().upper()
    if pod in DIRECT_POD_GROUPS:
        return pod
    if pod.startswith("KR"):
        return "KR"
    if pod.startswith("JP"):
        return "JP"
    if pod.startswith("RU"):
        return "RU"
    return "OTHER"


def resolve_lane(svc):
    svc = (svc or "").strip().upper()
    return SVC_TO_LANE.get(svc, svc)


def monday_of(dt):
    d = dt.date() if isinstance(dt, datetime) else dt
    return d - timedelta(days=d.weekday())


def status_for(pct):
    if pct is None:
        return "N/A"
    if abs(pct - 100.0) < 1e-6:
        return "FULL"
    if pct > 100.0:
        return "OVER"
    return "OK"


STATUS_ICON = {"OK": "\U0001F7E2", "OVER": "\U0001F534", "FULL": "\U0001F535", "N/A": "⚪"}

# The 5 POD views the report/dashboard offer -- matching the analyst's
# original Report.xlsx tab layout exactly: some destinations get their own
# tab, others are grouped together with their ALLO/Actual/% shown side by
# side. (sheet_name, dropdown_label, [sub-port names to show]). "TOTAL" is
# not one of PORT_GROUPS -- it's a synthetic column computed as KR+JP+RU
# combined, exactly like the original sheet's "TOTAL" column.
PORT_SHEET_GROUPS = [
    ("POD - VNSGN", "VNSGN", ["VNSGN"]),
    ("POD - HKG,CNXMN,CNSHK,TWKEL", "HKG, CNXMN, CNSHK, TWKEL", ["HKHKG", "CNXMN", "CNSHK", "TWKEL"]),
    ("POD - CNSHA", "CNSHA", ["CNSHA"]),
    ("POD - IDJKT", "IDJKT", ["IDJKT"]),
    ("POD - TOTAL,KR,JP,RU", "TOTAL, KR, JP, RU", ["TOTAL", "KR", "JP", "RU"]),
]


def get_port_cell(vessel_row, name):
    """Return the {allo, actual, pct, status} dict for one sub-port name on a
    vessel row, or None if that lane has nothing to show for it. "TOTAL" is
    computed on the fly as KR + JP + RU combined."""
    if name == "TOTAL":
        subs = [next((p for p in vessel_row["ports"] if p["port"] == s), None) for s in ("KR", "JP", "RU")]
        if not any(subs):
            return None
        allo_sum = sum((s["allo"] or 0) for s in subs if s)
        actual_sum = sum((s["actual"] or 0) for s in subs if s)
        pct = (actual_sum / allo_sum * 100) if allo_sum else None
        return {"port": "TOTAL", "allo": allo_sum if allo_sum else None, "actual": actual_sum, "pct": pct,
                "status": status_for(pct) if allo_sum else "N/A"}
    return next((p for p in vessel_row["ports"] if p["port"] == name), None)


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def build_analysis(bookings, sked_lookup, bsa):
    # Step 1: aggregate booking lines into one record per vessel sailing (VSL+VOY).
    sailings = {}
    for b in bookings:
        key = (b["vsl"], b["voy"])
        s = sailings.setdefault(key, {
            "svc_from_booking": None, "teu_by_port": defaultdict(float),
            "total_teu": 0.0, "total_weight_kg": 0.0, "etd_from_booking": None,
            "reefer_by_port": defaultdict(int), "total_reefer": 0,
        })
        if b["svc"] and not s["svc_from_booking"]:
            s["svc_from_booking"] = b["svc"]
        grp = pod_group(b["pod"])
        s["teu_by_port"][grp] += b["teu"]
        s["total_teu"] += b["teu"]
        s["total_weight_kg"] += b["weight_kg"]
        if b["reefer_qty"]:
            s["reefer_by_port"][grp] += b["reefer_qty"]
            s["total_reefer"] += b["reefer_qty"]
        if b["etd"] and not s["etd_from_booking"]:
            s["etd_from_booking"] = b["etd"]

    sailing_records = []
    for (vsl, voy), s in sailings.items():
        sked = sked_lookup.get((vsl, voy))
        svc_raw = (sked["service"] if sked else "") or s["svc_from_booking"] or ""
        lane = resolve_lane(svc_raw) if svc_raw else "UNMAPPED"
        vessel_name = sked["vessel_name"] if sked else vsl
        etd = (sked["etd"] if sked else None) or s["etd_from_booking"]
        sailing_records.append({
            "vsl": vsl, "voy": voy, "svc_raw": svc_raw, "lane": lane,
            "vessel_name": vessel_name, "etd": etd,
            "teu_by_port": dict(s["teu_by_port"]),
            "booked_teu": s["total_teu"],
            "booked_weight_ton": s["total_weight_kg"] / WEIGHT_DIVISOR,
            "reefer_by_port": dict(s["reefer_by_port"]),
            "total_reefer": s["total_reefer"],
            "slot_share": bool(sked and sked["slot_share"]),
            "slot_share_label": sked["slot_share_label"] if sked else "",
        })

    # Step 2: group sailings by ISO week (Mon-Sun) of ETD. Each vessel sailing
    # gets its own row -- compared against its lane's full BSA quota on its
    # own, never combined with other vessels of the same lane/week, since the
    # analysis is per vessel, not per SVC.
    weeks = defaultdict(list)
    for rec in sailing_records:
        week_monday = monday_of(rec["etd"]) if rec["etd"] else None
        weeks[week_monday].append(rec)

    week_blocks = []
    for week_monday in sorted(weeks.keys(), key=lambda d: (d is None, d)):
        if week_monday is None:
            week_label, week_range = "Unknown week (no ETD)", ""
        else:
            iso_week = week_monday.isocalendar()[1]
            week_sunday = week_monday + timedelta(days=6)
            week_label = f"Week {iso_week}"
            week_range = f"{week_monday.isoformat()} - {week_sunday.isoformat()}"

        recs = sorted(weeks[week_monday], key=lambda r: (r["lane"], r["etd"] or datetime.max, r["vessel_name"]))

        vessel_rows = []
        for r in recs:
            lane_bsa = bsa.get(r["lane"])
            teu_factor = SLOT_SHARE_TEU_FACTOR if r["slot_share"] else 1.0
            bsa_total_teu_full = lane_bsa["total_teu"] if lane_bsa else None
            bsa_total_teu = bsa_total_teu_full * teu_factor if bsa_total_teu_full is not None else None
            pct_teu = (r["booked_teu"] / bsa_total_teu * 100) if bsa_total_teu else None

            ports = []
            for p in PORT_GROUPS + (["OTHER"] if "OTHER" in r["teu_by_port"] else []):
                allo_full = lane_bsa["ports"].get(p) if lane_bsa else None
                allo = allo_full * teu_factor if allo_full is not None else None
                actual = r["teu_by_port"].get(p, 0.0)
                if allo is None and actual == 0.0:
                    continue
                pct = (actual / allo * 100) if allo else None
                ports.append({
                    "port": p, "allo": allo, "actual": actual, "pct": pct,
                    "status": status_for(pct) if allo else "N/A",
                })

            vessel_rows.append({
                "lane": r["lane"],
                "svc_raw": r["svc_raw"] or "UNMAPPED (no SVC/schedule match)",
                "vessel": f"{r['vessel_name']} ({r['vsl']} {r['voy']})",
                "etd": r["etd"],
                "bsa_full_teu": bsa_total_teu_full,
                "bsa_total_teu": bsa_total_teu,
                "booked_teu": r["booked_teu"],
                "pct_teu": pct_teu,
                "status": status_for(pct_teu),
                "booked_weight_ton": r["booked_weight_ton"],
                "ports": ports,
                "reefer_by_port": r["reefer_by_port"],
                "total_reefer": r["total_reefer"],
                "slot_share": r["slot_share"],
                "slot_share_notes": [r["slot_share_label"]] if r["slot_share"] else [],
            })

        week_blocks.append({"label": week_label, "range": week_range, "monday": week_monday, "rows": vessel_rows})

    return week_blocks


# ---------------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------------

def _style_header(ws, headers, header_fill, header_font, alt_fill=None, alt_cols=()):
    from openpyxl.styles import Alignment
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = alt_fill if (alt_fill and c in alt_cols) else header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def _merge_week_column(ws, start_row, end_row):
    from openpyxl.styles import Alignment
    if end_row > start_row:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
    ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="center", vertical="center")


def write_excel_report(week_blocks):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    WEEK_LABEL_FILL = PatternFill("solid", fgColor="D9D9D9")
    FILLS = {
        "OK": PatternFill("solid", fgColor="C6EFCE"),
        "OVER": PatternFill("solid", fgColor="FFC7CE"),
        "FULL": PatternFill("solid", fgColor="BDD7EE"),
        "N/A": None,
    }

    wb = openpyxl.Workbook()

    # ---- Main sheet: one row per vessel sailing, overall BSA/ALLO/LIFTING/% ----
    ws = wb.active
    ws.title = "Weekly Summary"
    headers = ["Week", "SVC", "Vessel / Voyage", "ETD", "BSA", "ALLO", "LIFTING", "%", "TEU Note", "Status",
               "Weight (ton)"]
    _style_header(ws, headers, HEADER_FILL, HEADER_FONT)

    for block in week_blocks:
        if not block["rows"]:
            continue
        start_row = ws.max_row + 1
        for vessel_row in block["rows"]:
            etd_str = vessel_row["etd"].strftime("%Y-%m-%d") if vessel_row["etd"] else "N/A"
            pct_str = f"{vessel_row['pct_teu']:.1f}%" if vessel_row["pct_teu"] is not None else "N/A"
            teu_note = f"USE {int(SLOT_SHARE_TEU_FACTOR * 100)}%" if vessel_row["slot_share"] else ""
            ws.append([
                block["label"], vessel_row["svc_raw"], vessel_row["vessel"], etd_str,
                vessel_row["bsa_full_teu"] if vessel_row["bsa_full_teu"] is not None else "N/A",
                vessel_row["bsa_total_teu"] if vessel_row["bsa_total_teu"] is not None else "N/A",
                round(vessel_row["booked_teu"], 2), pct_str, teu_note,
                STATUS_ICON[vessel_row["status"]] + " " + vessel_row["status"],
                round(vessel_row["booked_weight_ton"], 1),
            ])
            r = ws.max_row
            fill = FILLS.get(vessel_row["status"])
            if fill:
                for c in range(1, len(headers) + 1):
                    ws.cell(row=r, column=c).fill = fill
            if vessel_row["slot_share"]:
                ws.cell(row=r, column=9).font = Font(bold=True, color="0000FF")
        _merge_week_column(ws, start_row, ws.max_row)
        ws.cell(row=start_row, column=1).fill = WEEK_LABEL_FILL

    for i, w in enumerate([10, 10, 34, 12, 10, 10, 10, 10, 10, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # ---- 5 POD sheets, matching the analyst's original tab layout: some
    # destinations grouped together with their ALLO/Actual/% side by side ----
    GROUP_COLORS = ["2E75B6", "548235", "BF8F00", "7030A0"]  # blue, green, gold, purple -- cycled per sub-port
    for sheet_name, _label, sub_ports in PORT_SHEET_GROUPS:
        wsp = wb.create_sheet(sheet_name)
        headers_p = ["Week", "SVC", "Vessel / Voyage", "ETD", "BSA", "ALLO", "LIFTING", "%"]
        for sp in sub_ports:
            headers_p += [f"{sp} ALLO", f"{sp} Actual", f"{sp} %"]
        col_fills = {}
        for gi, sp in enumerate(sub_ports):
            color = PatternFill("solid", fgColor=GROUP_COLORS[gi % len(GROUP_COLORS)])
            base_col = 9 + gi * 3
            col_fills[base_col] = color
            col_fills[base_col + 1] = color
            col_fills[base_col + 2] = color
        wsp.append(headers_p)
        for c in range(1, len(headers_p) + 1):
            cell = wsp.cell(row=1, column=c)
            cell.font = HEADER_FONT
            cell.fill = col_fills.get(c, HEADER_FILL)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        wsp.freeze_panes = "A2"

        for block in week_blocks:
            rows_for_group = [
                vr for vr in block["rows"] if any(get_port_cell(vr, sp) for sp in sub_ports)
            ]
            if not rows_for_group:
                continue
            start_row = wsp.max_row + 1
            for vessel_row in rows_for_group:
                etd_str = vessel_row["etd"].strftime("%Y-%m-%d") if vessel_row["etd"] else "N/A"
                pct_str = f"{vessel_row['pct_teu']:.1f}%" if vessel_row["pct_teu"] is not None else "N/A"
                row = [
                    block["label"], vessel_row["svc_raw"], vessel_row["vessel"], etd_str,
                    vessel_row["bsa_full_teu"] if vessel_row["bsa_full_teu"] is not None else "N/A",
                    vessel_row["bsa_total_teu"] if vessel_row["bsa_total_teu"] is not None else "N/A",
                    round(vessel_row["booked_teu"], 2), pct_str,
                ]
                port_infos = [get_port_cell(vessel_row, sp) for sp in sub_ports]
                for info in port_infos:
                    if info is None:
                        row += ["N/A", "-", "N/A"]
                    else:
                        port_pct_str = f"{info['pct']:.1f}%" if info["pct"] is not None else "N/A"
                        row += [info["allo"] if info["allo"] is not None else "N/A", round(info["actual"], 2), port_pct_str]
                wsp.append(row)
                r = wsp.max_row
                overall_fill = FILLS.get(vessel_row["status"])
                if overall_fill:
                    for c in (5, 6, 7, 8):
                        wsp.cell(row=r, column=c).fill = overall_fill
                for gi, info in enumerate(port_infos):
                    if info is None:
                        continue
                    fill = FILLS.get(info["status"])
                    if fill:
                        base_col = 9 + gi * 3
                        for c in (base_col, base_col + 1, base_col + 2):
                            wsp.cell(row=r, column=c).fill = fill
                if vessel_row["slot_share"]:
                    wsp.cell(row=r, column=9).font = Font(bold=True, color="0000FF")
            _merge_week_column(wsp, start_row, wsp.max_row)
            wsp.cell(row=start_row, column=1).fill = WEEK_LABEL_FILL

        widths = [10, 10, 34, 12, 10, 10, 10, 10] + [12, 12, 10] * len(sub_ports)
        for i, w in enumerate(widths, start=1):
            wsp.column_dimensions[get_column_letter(i)].width = w
        wsp.auto_filter.ref = f"A1:{get_column_letter(len(headers_p))}1"

    # ---- Reefer (Plug) sheet: every sailing that carries at least one
    # reefer container, broken out by destination port group ----
    ws3 = wb.create_sheet("Reefer")
    reefer_ports = [p for p in PORT_GROUPS if p in {
        p for block in week_blocks for row in block["rows"] for p in row["reefer_by_port"]
    }] + (["OTHER"] if any("OTHER" in row["reefer_by_port"] for block in week_blocks for row in block["rows"]) else [])
    headers3 = ["Week", "SVC", "Vessel / Voyage", "ETD", "Total Reefer Qty"] + [f"{p} Reefer Qty" for p in reefer_ports]
    ws3.append(headers3)
    for c in range(1, len(headers3) + 1):
        cell = ws3.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws3.freeze_panes = "A2"
    port_totals = defaultdict(int)
    grand_total = 0
    for block in week_blocks:
        for vessel_row in block["rows"]:
            if not vessel_row["total_reefer"]:
                continue
            etd_str = vessel_row["etd"].strftime("%Y-%m-%d") if vessel_row["etd"] else "N/A"
            row = [block["label"], vessel_row["svc_raw"], vessel_row["vessel"], etd_str, vessel_row["total_reefer"]]
            for p in reefer_ports:
                qty = vessel_row["reefer_by_port"].get(p, 0)
                row.append(qty)
                port_totals[p] += qty
            ws3.append(row)
            grand_total += vessel_row["total_reefer"]
    ws3.append(["", "TOTAL", "", "", grand_total] + [port_totals.get(p, 0) for p in reefer_ports])
    for c in range(1, len(headers3) + 1):
        ws3.cell(row=ws3.max_row, column=c).font = Font(bold=True)
    for i, w in enumerate([10, 10, 34, 12, 14] + [14] * len(reefer_ports), start=1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(headers3))}1"

    # ---- Alerts sheet ----
    ws2 = wb.create_sheet("Alerts (OVER, FULL)")
    ws2.append(["Week", "SVC", "Vessel / Voyage", "BSA", "ALLO", "LIFTING", "%", "TEU Note", "Status", "Excess TEU"])
    for c in range(1, 11):
        cell = ws2.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    any_alert = False
    for block in week_blocks:
        for vessel_row in block["rows"]:
            if vessel_row["status"] not in ("OVER", "FULL"):
                continue
            any_alert = True
            excess = (vessel_row["booked_teu"] - vessel_row["bsa_total_teu"]) if vessel_row["bsa_total_teu"] else None
            teu_note = f"USE {int(SLOT_SHARE_TEU_FACTOR * 100)}%" if vessel_row["slot_share"] else ""
            ws2.append([
                block["label"], vessel_row["svc_raw"], vessel_row["vessel"],
                vessel_row["bsa_full_teu"], vessel_row["bsa_total_teu"], round(vessel_row["booked_teu"], 2),
                f"{vessel_row['pct_teu']:.1f}%" if vessel_row["pct_teu"] is not None else "N/A",
                teu_note,
                STATUS_ICON[vessel_row["status"]] + " " + vessel_row["status"],
                round(excess, 2) if excess is not None else "N/A",
            ])
            fill = FILLS.get(vessel_row["status"])
            if fill:
                for c in range(1, 11):
                    ws2.cell(row=ws2.max_row, column=c).fill = fill
            if vessel_row["slot_share"]:
                ws2.cell(row=ws2.max_row, column=8).font = Font(bold=True, color="0000FF")
    if not any_alert:
        ws2.append(["No vessel sailing is OVER or FULL this run."])
    for i, w in enumerate([10, 10, 34, 10, 10, 10, 10, 10, 12, 12], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    wb.save(EXCEL_OUTPUT)


# ---------------------------------------------------------------------------
# HTML dashboard
# ---------------------------------------------------------------------------

STATUS_COLOR = {"OK": "#1e7e34", "OVER": "#c0392b", "FULL": "#1b5fad", "N/A": "#888888"}
STATUS_BG = {"OK": "#e6f6ea", "OVER": "#fdecea", "FULL": "#e8f0fb", "N/A": "#f2f2f2"}


def _dashboard_json(week_blocks):
    data = []
    for block in week_blocks:
        rows = []
        for vr in block["rows"]:
            rows.append({
                "svc": vr["svc_raw"],
                "vessel": vr["vessel"],
                "etd": vr["etd"].strftime("%Y-%m-%d") if vr["etd"] else None,
                "bsaFull": vr["bsa_full_teu"],
                "bsaAllo": vr["bsa_total_teu"],
                "lifting": round(vr["booked_teu"], 2),
                "pctTeu": round(vr["pct_teu"], 1) if vr["pct_teu"] is not None else None,
                "status": vr["status"],
                "ports": [
                    {"port": p["port"], "allo": p["allo"], "actual": round(p["actual"], 2),
                     "pct": round(p["pct"], 1) if p["pct"] is not None else None, "status": p["status"]}
                    for p in vr["ports"]
                ],
                "notes": vr["slot_share_notes"],
            })
        data.append({"label": block["label"], "range": block["range"], "rows": rows})
    return data


def _reefer_json(week_blocks):
    reefer_ports = sorted({p for block in week_blocks for row in block["rows"] for p in row["reefer_by_port"]})
    rows = []
    for block in week_blocks:
        for vr in block["rows"]:
            if not vr["total_reefer"]:
                continue
            rows.append({
                "week": block["label"], "svc": vr["svc_raw"], "vessel": vr["vessel"],
                "etd": vr["etd"].strftime("%Y-%m-%d") if vr["etd"] else None,
                "total": vr["total_reefer"],
                "ports": {p: vr["reefer_by_port"].get(p, 0) for p in reefer_ports},
            })
    return {"ports": reefer_ports, "rows": rows}


def write_html_dashboard(week_blocks, generated_at):
    payload = _dashboard_json(week_blocks)
    reefer_payload = _reefer_json(week_blocks)
    port_groups_js = [{"key": name, "label": label, "ports": ports} for name, label, ports in PORT_SHEET_GROUPS]
    total_rows = sum(len(b["rows"]) for b in payload)
    counts = defaultdict(int)
    for b in payload:
        for row in b["rows"]:
            counts[row["status"]] += 1

    html = r"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<title>BSA Utilization Report</title>
<style>
  :root {
    --ok: #15803d; --ok-bg: #ecfdf3; --ok-border: #bbf0cd;
    --over: #c0392b; --over-bg: #fef2f1; --over-border: #f8c9c4;
    --full: #1d4ed8; --full-bg: #eef3ff; --full-border: #c7d7fb;
    --na: #94a3b8; --na-bg: #f4f6f8; --na-border: #e4e9ef;
    --ink: #101828; --muted: #6b7887; --border: #e7ebf0; --band: #f9fafb;
    --accent: #2563eb; --accent-soft: #eef4ff;
    --row-line: #eef0f3; --row-alt: #f8f9fb; --row-hover: #f1f5fb;
    --shadow: 0 1px 2px rgba(16,24,40,.04), 0 2px 8px rgba(16,24,40,.05);
  }
  * { box-sizing: border-box; }
  html { -webkit-text-size-adjust:100%; }
  body { margin:0; background:#f4f6f9; color:var(--ink); font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif; -webkit-font-smoothing:antialiased; }
  header { background:#fff; border-bottom:1px solid var(--border); padding:18px 28px; }
  header .header-inner { max-width:1600px; margin:0 auto; display:flex; align-items:center; justify-content:space-between; gap:14px; flex-wrap:wrap; }
  header .header-left { display:flex; align-items:center; gap:14px; }
  header .header-dot { font-size:20px; line-height:1; flex:none; }
  header h1 { margin:0; font-size:20px; font-weight:700; color:var(--ink); letter-spacing:-.01em; }
  header p { margin:2px 0 0; color:var(--muted); font-size:13px; }
  header .header-right { display:flex; align-items:center; gap:16px; }
  header .header-logo { height:36px; width:auto; object-fit:contain; border-left:1px solid var(--border); padding-left:16px; }
  header .header-clock { text-align:right; }
  header .header-clock .clock-time { font-size:16px; font-weight:700; color:var(--ink); font-variant-numeric:tabular-nums; letter-spacing:.02em; }
  header .header-clock .clock-date { font-size:11.5px; color:var(--muted); font-weight:500; margin-top:1px; }
  .wrap { max-width:1600px; margin:0 auto; padding:24px 28px 40px; }
  .kpis { display:flex; gap:16px; flex-wrap:wrap; margin-bottom:22px; }
  .kpi { position:relative; flex:1; min-width:170px; background:#fff; border:1px solid var(--border); border-radius:16px; padding:18px 20px 16px; box-shadow:var(--shadow); overflow:hidden; }
  .kpi::before { content:''; position:absolute; top:0; left:0; right:0; height:5px; background:linear-gradient(90deg, var(--c1), var(--c2)); }
  .kpi::after { content:''; position:absolute; width:96px; height:96px; border-radius:50%; top:-44px; right:-32px; background:var(--c1); opacity:.09; }
  .kpi-icon { position:relative; width:34px; height:34px; border-radius:10px; display:flex; align-items:center; justify-content:center; font-size:16px; background:var(--c1-soft); margin-bottom:12px; }
  .kpi .num { position:relative; font-size:28px; font-weight:700; line-height:1.2; }
  .kpi .lbl { position:relative; font-size:12.5px; color:var(--muted); margin-top:3px; font-weight:600; }
  .kpi.total { --c1:#2563eb; --c2:#7dabfb; --c1-soft:var(--accent-soft); }
  .kpi.ok { --c1:#15803d; --c2:#5fd68a; --c1-soft:var(--ok-bg); } .kpi.ok .num { color:var(--ok); }
  .kpi.full { --c1:#1d4ed8; --c2:#7dabfb; --c1-soft:var(--full-bg); } .kpi.full .num { color:var(--full); }
  .kpi.over { --c1:#c0392b; --c2:#f18f85; --c1-soft:var(--over-bg); } .kpi.over .num { color:var(--over); }
  .controls { display:flex; gap:12px; align-items:center; margin-bottom:20px; flex-wrap:wrap; background:#fff; border:1px solid var(--border); border-radius:12px; padding:12px 16px; box-shadow:var(--shadow); }
  .controls label { font-size:13px; color:var(--muted); font-weight:500; }
  select { padding:8px 12px; border:1px solid var(--border); border-radius:8px; font-size:13.5px; background:#fff; color:var(--ink); cursor:pointer; }
  select:focus { outline:none; border-color:var(--accent); box-shadow:0 0 0 3px var(--accent-soft); }
  .week { background:#fff; border:1px solid var(--border); border-radius:14px; margin-bottom:18px; box-shadow:var(--shadow); overflow:hidden; }
  .week-head { padding:14px 20px; background:var(--band); border-bottom:1px solid var(--border); font-weight:700; font-size:14.5px; display:flex; justify-content:space-between; align-items:center; }
  .week-head .range { font-weight:500; color:var(--muted); font-size:12.5px; }
  .table-scroll { overflow-x:auto; -webkit-overflow-scrolling:touch; }
  table { width:100%; border-collapse:collapse; font-size:13.5px; table-layout:auto; }
  th, td { padding:13px 16px; border-bottom:1px solid var(--row-line); text-align:left; white-space:nowrap; }
  td.wrap-cell { white-space:normal; min-width:180px; }
  td { color:var(--ink); font-weight:500; }
  th { color:var(--muted); font-weight:600; font-size:11px; text-transform:uppercase; letter-spacing:.05em; background:var(--band); border-bottom:1px solid var(--border); }
  table.compact { width:100%; table-layout:fixed; }
  table.compact th, table.compact td { padding:10px 8px; font-size:12.5px; white-space:normal; word-break:break-word; }
  table.compact th { font-size:10px; line-height:1.3; }
  table.compact td.wrap-cell { min-width:0; }
  tbody tr:nth-child(even) { background:var(--row-alt); }
  tbody tr:hover { background:var(--row-hover); }
  tbody tr:last-child td { border-bottom:none; }
  .pill { display:inline-flex; align-items:center; gap:5px; padding:4px 11px; border-radius:999px; font-weight:600; font-size:12.5px; border:1px solid transparent; }
  .pill.OK { color:var(--ok); background:var(--ok-bg); border-color:var(--ok-border); }
  .pill.OVER { color:var(--over); background:var(--over-bg); border-color:var(--over-border); }
  .pill.FULL { color:var(--full); background:var(--full-bg); border-color:var(--full-border); }
  .pill.N-A { color:var(--na); background:var(--na-bg); border-color:var(--na-border); }
  .pct-text { font-weight:700; }
  .pct-text.OK { color:var(--ok); } .pct-text.OVER { color:var(--over); } .pct-text.FULL { color:var(--full); } .pct-text.N-A { color:var(--na); }
  .note { color:var(--full); font-size:12px; margin-top:4px; font-weight:500; }
  .empty { padding:32px; text-align:center; color:var(--muted); background:#fff; border:1px solid var(--border); border-radius:14px; }
  footer { text-align:center; color:var(--muted); font-size:12px; padding:24px; }
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <div class="header-left">
      <span class="header-dot">&#x1F6A2;</span>
      <div>
        <h1>BSA Utilization Report</h1>
        <p>Booking vs BSA - VNSGN - HKHKG - CNXMN - CNSHK - TWKEL - CNSHA - KRPUS - IDJKT -</p>
      </div>
    </div>
    <div class="header-right">
      <div class="header-clock">
        <div class="clock-time" id="clockTime">--:--:--</div>
        <div class="clock-date" id="clockDate">Loading...</div>
      </div>
      <img class="header-logo" src="logo.png" alt="Company logo">
    </div>
  </div>
</header>
<div class="wrap">
  <div class="kpis">
    <div class="kpi total"><div class="kpi-icon">&#x1F6A2;</div><div class="num">__TOTAL_LANES__</div><div class="lbl">Vessel sailings</div></div>
    <div class="kpi ok"><div class="kpi-icon">&#x1F7E2;</div><div class="num">__COUNT_OK__</div><div class="lbl">OK</div></div>
    <div class="kpi full"><div class="kpi-icon">&#x1F535;</div><div class="num">__COUNT_FULL__</div><div class="lbl">100% (Full)</div></div>
    <div class="kpi over"><div class="kpi-icon">&#x1F534;</div><div class="num">__COUNT_OVER__</div><div class="lbl">OVER</div></div>
  </div>
  <div class="controls">
    <label for="portSelect">View by destination Port (POD):</label>
    <select id="portSelect">
      <option value="__ALL__">All ports (TEU total)</option>
      __PORT_OPTIONS__
    </select>
    <label for="statusFilter">Status:</label>
    <select id="statusFilter">
      <option value="">All</option>
      <option value="OK">OK</option>
      <option value="FULL">Full (100%)</option>
      <option value="OVER">Over</option>
    </select>
  </div>
  <div id="weeks"></div>
  <div id="reeferSection"></div>
</div>
<footer>Extraction: openpyxl / xlrd / pdfplumber. Analysis: deterministic Python (no AI). Weeks run Monday&ndash;Sunday.</footer>

<script id="dashboard-data" type="application/json">__DATA_JSON__</script>
<script id="port-groups-data" type="application/json">__PORT_GROUPS_JSON__</script>
<script id="reefer-data" type="application/json">__REEFER_JSON__</script>
<script>
const DATA = JSON.parse(document.getElementById('dashboard-data').textContent);
const PORT_GROUPS = JSON.parse(document.getElementById('port-groups-data').textContent);
const REEFER = JSON.parse(document.getElementById('reefer-data').textContent);

function pill(status, text) {
  const cls = status.replace('/', '-');
  return `<span class="pill ${cls}">${text}</span>`;
}
function pctSpan(pct, status) {
  const cls = status.replace('/', '-');
  const text = (pct === null || pct === undefined) ? 'N/A' : pct.toFixed(1) + '%';
  return `<span class="pct-text ${cls}">${text}</span>`;
}
function statusIcon(s) {
  return {OK: '\u{1F7E2}', OVER: '\u{1F534}', FULL: '\u{1F535}', 'N/A': '⚪'}[s] || '';
}
function getPortCell(row, name) {
  if (name === 'TOTAL') {
    const subs = ['KR', 'JP', 'RU'].map(s => row.ports.find(x => x.port === s)).filter(Boolean);
    if (!subs.length) return null;
    const allo = subs.reduce((a, s) => a + (s.allo || 0), 0);
    const actual = subs.reduce((a, s) => a + (s.actual || 0), 0);
    const pct = allo ? (actual / allo * 100) : null;
    const status = allo ? (pct === 100 ? 'FULL' : (pct > 100 ? 'OVER' : 'OK')) : 'N/A';
    return {port: 'TOTAL', allo: allo || null, actual, pct, status};
  }
  return row.ports.find(x => x.port === name) || null;
}

const GROUP_COLORS = [
  {head: '#0EA5E9', soft: '#f2faff', softAlt: '#e0f2fe'},
  {head: '#16A34A', soft: '#f3fbf6', softAlt: '#dcf7e6'},
  {head: '#DB2777', soft: '#fef4f9', softAlt: '#fbe0ee'},
  {head: '#7C3AED', soft: '#f9f5fe', softAlt: '#eee0fc'},
];

function render() {
  const groupKey = document.getElementById('portSelect').value;
  const group = PORT_GROUPS.find(g => g.key === groupKey);
  const statusFilter = document.getElementById('statusFilter').value;
  const root = document.getElementById('weeks');
  root.innerHTML = '';

  // Builds the style for one cell inside a coloured port-group block: a
  // small white gap + rounded corner between groups, and (on the header or
  // the last body row) a rounded outer edge so each group reads as a
  // soft rounded "chip" rather than a sharp rectangle.
  function groupCellStyle(bg, isFirstInGroup, isLastInGroup, isLastGroup, roundTop, roundBottom) {
    let s = `background:${bg};`;
    if (!isLastGroup && isLastInGroup) s += 'border-right:4px solid #fff;';
    const tl = roundTop && isFirstInGroup ? '10px' : '0';
    const tr = roundTop && isLastInGroup ? '10px' : '0';
    const br = roundBottom && isLastInGroup ? '10px' : '0';
    const bl = roundBottom && isFirstInGroup ? '10px' : '0';
    s += `border-radius:${tl} ${tr} ${br} ${bl};`;
    return s;
  }

  DATA.forEach(block => {
    const renderedRows = [];
    let rowIndex = 0;
    block.rows.forEach(row => {
      if (statusFilter && row.status !== statusFilter) return;
      renderedRows.push({row, portCells: null});
      if (group) {
        const cells = group.ports.map(p => getPortCell(row, p));
        if (!cells.some(Boolean)) { renderedRows.pop(); return; }
        renderedRows[renderedRows.length - 1].portCells = cells;
        renderedRows[renderedRows.length - 1].shade = rowIndex % 2 === 0 ? 'soft' : 'softAlt';
      }
      rowIndex++;
    });
    if (!renderedRows.length) return;

    // In a specific-port view, drop the overall BSA/ALLO/LIFTING/% columns
    // (already visible in "All ports") and collapse each port's Allo+Actual
    // into one cell, so the whole picture fits without horizontal scrolling.
    const rowsHtml = renderedRows.map((entry, ri) => {
      const {row, portCells, shade} = entry;
      const isLastRow = ri === renderedRows.length - 1;
      let portCellsHtml = '';
      let baseCellsHtml;
      if (group) {
        portCellsHtml = group.ports.map((p, i) => {
          const c = portCells[i];
          const bg = GROUP_COLORS[i % GROUP_COLORS.length][shade];
          const isLast = i === group.ports.length - 1;
          const sAllo = `style="${groupCellStyle(bg, true, false, isLast, false, isLastRow)}"`;
          const sPct = `style="${groupCellStyle(bg, false, true, isLast, false, isLastRow)}"`;
          if (!c) return `<td ${sAllo}>N/A</td><td ${sPct}>N/A</td>`;
          const allo = c.allo === null || c.allo === undefined ? 'N/A' : c.allo;
          return `<td ${sAllo}>${allo} / ${c.actual}</td><td ${sPct}>${pctSpan(c.pct, c.status)}</td>`;
        }).join('');
        baseCellsHtml = `
        <td class="wrap-cell">${row.vessel}${row.notes.length ? `<div class="note">USE 60% (slot-share: ${row.notes.join('; ')})</div>` : ''}</td>
        <td>${row.etd || 'N/A'}</td>
        <td>${pill(row.status, statusIcon(row.status) + ' ' + row.status)}</td>`;
      } else {
        baseCellsHtml = `
        <td>${row.svc}</td>
        <td class="wrap-cell">${row.vessel}${row.notes.length ? `<div class="note">USE 60% (slot-share: ${row.notes.join('; ')})</div>` : ''}</td>
        <td>${row.etd || 'N/A'}</td>
        <td>${row.bsaFull === null || row.bsaFull === undefined ? 'N/A' : row.bsaFull}</td>
        <td>${row.bsaAllo === null || row.bsaAllo === undefined ? 'N/A' : row.bsaAllo}</td>
        <td>${row.lifting}</td>
        <td>${pctSpan(row.pctTeu, row.status)}</td>
        <td>${pill(row.status, statusIcon(row.status) + ' ' + row.status)}</td>`;
      }
      return `<tr>${baseCellsHtml}${portCellsHtml}</tr>`;
    }).join('');

    const portHeadHtml = !group ? '' :
      group.ports.map((p, i) => {
        const c = GROUP_COLORS[i % GROUP_COLORS.length];
        const isLast = i === group.ports.length - 1;
        const sFirst = `style="color:#fff;${groupCellStyle(c.head, true, false, isLast, true, false)}"`;
        const sLast = `style="color:#fff;${groupCellStyle(c.head, false, true, isLast, true, false)}"`;
        return `<th ${sFirst}>${p} Allo / Actual</th><th ${sLast}>${p} %</th>`;
      }).join('');
    const theadHtml = !group
      ? `<th>SVC</th><th>Vessel / Voyage</th><th>ETD</th><th>BSA</th><th>ALLO</th><th>LIFTING</th><th>%</th><th>Status</th>`
      : `<th>Vessel / Voyage</th><th>ETD</th><th>Status</th>${portHeadHtml}`;
    root.insertAdjacentHTML('beforeend', `
      <div class="week">
        <div class="week-head"><span>${block.label}</span><span class="range">${block.range}</span></div>
        <div class="table-scroll">
        <table class="${group ? 'compact' : ''}">
          <thead><tr>${theadHtml}</tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
        </div>
      </div>`);
  });
  if (!root.innerHTML) {
    root.innerHTML = '<div class="empty">No sailings match this filter.</div>';
  }
}

document.getElementById('portSelect').addEventListener('change', render);
document.getElementById('statusFilter').addEventListener('change', render);
render();

function renderReefer() {
  const root = document.getElementById('reeferSection');
  if (!REEFER.rows.length) {
    root.innerHTML = '';
    return;
  }
  const portHeads = REEFER.ports.map(p => `<th>${p} Qty</th>`).join('');
  const rowsHtml = REEFER.rows.map(r => {
    const portCells = REEFER.ports.map(p => `<td>${r.ports[p] || 0}</td>`).join('');
    return `<tr>
      <td>${r.week}</td>
      <td>${r.svc}</td>
      <td class="wrap-cell">${r.vessel}</td>
      <td>${r.etd || 'N/A'}</td>
      <td><strong>${r.total}</strong></td>
      ${portCells}
    </tr>`;
  }).join('');
  root.innerHTML = `
    <div class="week">
      <div class="week-head"><span>&#x2744;&#xFE0F; Reefer (Plug) Bookings</span><span class="range">Every sailing carrying at least one reefer container</span></div>
      <div class="table-scroll">
      <table>
        <thead><tr>
          <th>Week</th><th>SVC</th><th>Vessel / Voyage</th><th>ETD</th><th>Total Reefer Qty</th>${portHeads}
        </tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      </div>
    </div>`;
}
renderReefer();

function tickClock() {
  const now = new Date();
  const timeFmt = new Intl.DateTimeFormat('en-US', {hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false});
  const dateFmt = new Intl.DateTimeFormat('en-US', {weekday: 'short', year: 'numeric', month: 'short', day: '2-digit'});
  document.getElementById('clockTime').textContent = timeFmt.format(now);
  document.getElementById('clockDate').textContent = dateFmt.format(now);
}
tickClock();
setInterval(tickClock, 1000);
</script>
</body>
</html>
"""
    port_options = "\n      ".join(f'<option value="{g["key"]}">{g["label"]}</option>' for g in port_groups_js)
    html = html.replace("__GENERATED_AT__", generated_at)
    html = html.replace("__TOTAL_LANES__", str(total_rows))
    html = html.replace("__COUNT_OK__", str(counts["OK"]))
    html = html.replace("__COUNT_FULL__", str(counts["FULL"]))
    html = html.replace("__COUNT_OVER__", str(counts["OVER"]))
    html = html.replace("__PORT_OPTIONS__", port_options)
    html = html.replace("__DATA_JSON__", json.dumps(payload, ensure_ascii=False))
    html = html.replace("__PORT_GROUPS_JSON__", json.dumps(port_groups_js, ensure_ascii=False))
    html = html.replace("__REEFER_JSON__", json.dumps(reefer_payload, ensure_ascii=False))

    with open(HTML_OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)
    with open(INDEX_OUTPUT, "w", encoding="utf-8") as f:
        f.write(html)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    bsa_file = find_bsa_file()
    sked_file = find_sked_file()
    master_file = find_master_booking_file()
    print(f"BSA file:     {os.path.basename(bsa_file)}")
    print(f"SKED file:    {os.path.basename(sked_file)}")
    print(f"Master file:  {os.path.basename(master_file)}")

    extract_pdf_texts()  # no-op today; picks up any future PDF drops

    bsa = parse_bsa(bsa_file)
    sked_lookup = parse_sked(sked_file)
    bookings = parse_bookings(master_file)
    print(f"Parsed {len(bookings)} booking lines, {len(sked_lookup)} scheduled sailings, {len(bsa)} BSA lanes.")

    week_blocks = build_analysis(bookings, sked_lookup, bsa)

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    write_excel_report(week_blocks)
    write_html_dashboard(week_blocks, generated_at)

    total_rows = sum(len(b["rows"]) for b in week_blocks)
    over = sum(1 for b in week_blocks for r in b["rows"] if r["status"] == "OVER")
    full = sum(1 for b in week_blocks for r in b["rows"] if r["status"] == "FULL")
    print(f"\nWrote {total_rows} vessel-sailing rows across {len(week_blocks)} weeks.")
    print(f"OVER: {over}   FULL: {full}   OK: {total_rows - over - full}")
    print(f"Excel report:    {EXCEL_OUTPUT}")
    print(f"HTML dashboard:  {HTML_OUTPUT}")
    print(f"                 {INDEX_OUTPUT} (same content, for GitHub Pages)")


if __name__ == "__main__":
    main()
