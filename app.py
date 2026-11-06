import io
import hashlib
import hmac
import json
import os
import re
import secrets
import smtplib
import sqlite3
import unicodedata
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import parse_qs
from zoneinfo import ZoneInfo
from xml.sax.saxutils import escape as xml_escape
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
TEMPLATE_PATHS = {
    "de": BASE_DIR / "templates" / "Tagesrapport-DE.xlsx",
    "fr": BASE_DIR / "templates" / "Tagesrapport-FR.xlsx",
    "it": BASE_DIR / "templates" / "Tagesrapport-IT.xlsx",
}
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data"))
DB_PATH = DATA_DIR / "tagesrapporte.db"
MAX_BODY = 2 * 1024 * 1024
PROJECT_LEADERS = [
    "Imer Abazi", "Fidan Salihi", "Damiano Giarrusso", "Megzon Morina",
    "Sandro Sargenti", "Armando Baldacci", "Nico Gavoldi", "Remo Brunner",
]
WORK_PRICES = {
    "Projektleiter": 110, "Vorarbeiter": 80, "Facharbeiter": 72,
    "Schadenaufnahme": 110, "Leckortung": 127, "Aufbau": 98,
    "Kontrollmessung": 98, "Teilabbau": 98, "Abbau": 98,
    "Löcher öffnen/schliessen": 98, "Putz/Tapeten entfernen": 98, "Trocknungsarbeit": 98,
}
VEHICLE_PRICES = {
    "PW": {"Kilometer": 1.90, "Pauschale": 35},
    "Bus": {"Kilometer": 2.15, "Pauschale": 50},
    "Kein Fahrzeug": {"Kilometer": 0, "Pauschale": 0},
}
MATERIAL_PRICES = {
    "Stark alkalischer Reiniger": 22.5, "Neutralreiniger": 22.5,
    "Glasreiniger": 13, "Spezialreiniger": 30, "Sanitärreiniger": 17.5,
    "Teppichreiniger": 22.5, "Langezeit Konservierungsöl": 27,
    "Geruchsregulator": 35, "Endoskopiekamera": 30, "Wärmebildkamera": 50,
    "Karton Schachteln": 6, "Abdeckvliess (Floorliner)": 5,
    "Abdeck-Baufolie": 1.6, "Luftpolsterfolie": 0, "Färbemittel Fluoreszin": 5,
}


def utc_now():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def db_connection():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS reports (
            id TEXT PRIMARY KEY,
            title TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Entwurf',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            payload TEXT NOT NULL
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS employees (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL UNIQUE COLLATE NOCASE,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS project_leaders (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL UNIQUE COLLATE NOCASE,
            email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            password_hash TEXT,
            password_salt TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS login_sessions (
            token_hash TEXT PRIMARY KEY,
            project_leader_id TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            FOREIGN KEY(project_leader_id) REFERENCES project_leaders(id)
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS office_workers (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE COLLATE NOCASE,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )
        """
    )
    for name in PROJECT_LEADERS:
        connection.execute(
            "INSERT OR IGNORE INTO project_leaders (id, name, email, updated_at) VALUES (?, ?, ?, ?)",
            (str(uuid.uuid4()), name, leader_email(name), utc_now()),
        )
    connection.commit()
    return connection


def ascii_slug(value):
    normalized = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "-", normalized.lower()).strip("-") or "unbekannt"


def leader_email(name):
    parts = ascii_slug(name).split("-")
    local = f"{parts[0]}.{parts[-1]}" if len(parts) > 1 else parts[0]
    return f"{local}@polygongroup.com"


def hash_password(password, salt=None):
    salt_bytes = bytes.fromhex(salt) if salt else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt_bytes, 310_000)
    return digest.hex(), salt_bytes.hex()


def project_leader_session(environ):
    header = environ.get("HTTP_AUTHORIZATION", "")
    if not header.startswith("Bearer "):
        return None
    token = header[7:].strip()
    if not token:
        return None
    token_hash = hashlib.sha256(token.encode("utf-8")).hexdigest()
    with db_connection() as connection:
        row = connection.execute(
            """SELECT p.id, p.name, p.email, s.expires_at
               FROM login_sessions s JOIN project_leaders p ON p.id = s.project_leader_id
               WHERE s.token_hash = ?""",
            (token_hash,),
        ).fetchone()
        if not row:
            return None
        if datetime.fromisoformat(row["expires_at"]) <= datetime.now(timezone.utc):
            connection.execute("DELETE FROM login_sessions WHERE token_hash = ?", (token_hash,))
            connection.commit()
            return None
    return dict(row)


def auth_context(environ):
    leader = project_leader_session(environ)
    if leader:
        return {"kind": "leader", **leader}
    expected = os.environ.get("APP_PASSWORD", "").strip()
    provided = environ.get("HTTP_X_APP_KEY", "")
    if (expected and hmac.compare_digest(provided, expected)) or not expected:
        return {"kind": "admin"}
    return None


def json_response(start_response, payload, status="200 OK", headers=None):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    final_headers = [("Content-Type", "application/json; charset=utf-8"), ("Content-Length", str(len(body)))]
    final_headers.extend(headers or [])
    start_response(status, final_headers)
    return [body]


def text_response(start_response, body, content_type, status="200 OK", headers=None):
    encoded = body if isinstance(body, bytes) else body.encode("utf-8")
    final_headers = [("Content-Type", content_type), ("Content-Length", str(len(encoded)))]
    final_headers.extend(headers or [])
    start_response(status, final_headers)
    return [encoded]


def read_json(environ):
    try:
        length = int(environ.get("CONTENT_LENGTH") or 0)
    except ValueError:
        length = 0
    if length <= 0 or length > MAX_BODY:
        raise ValueError("Ungültige Anfragegrösse")
    raw = environ["wsgi.input"].read(length)
    return json.loads(raw.decode("utf-8"))


def safe_title(data):
    project_number = str(data.get("projectNumber") or "").strip()
    project_name = str(data.get("projectName") or (data.get("location") or {}).get("name") or "Tagesrapport").strip()
    date = data.get("date") or datetime.now().date().isoformat()
    project = " - ".join(part for part in (project_number, project_name) if part)
    return f"{project} – {date}"[:160]


def normalize_report(data):
    if not isinstance(data, dict):
        raise ValueError("Rapportdaten fehlen")
    required = ["date", "reporterName", "region", "projectLead", "damageType", "projectNumber", "projectName", "projectStatus"]
    missing = [field for field in required if not str(data.get(field, "")).strip()]
    if not str((data.get("location") or {}).get("city") or "").strip():
        missing.append("Ort")
    costs = data.get("costs") or {}
    if costs.get("inspectorEmail") and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", str(costs["inspectorEmail"])):
        missing.append("gültige E-Mail SI")
    if missing:
        raise ValueError("Bitte Pflichtfelder ausfüllen: " + ", ".join(missing))
    data["areas"] = [area for area in data.get("areas", []) if area in {"fh", "gb", "tr"}]
    data["projectStatus"] = data.get("projectStatus") if data.get("projectStatus") in {"Offen", "In Arbeit", "Pausiert", "Abgeschlossen"} else "Offen"
    data["thirdParty"] = list(data.get("thirdParty") or [])[:6]
    return data


def validate_submission(report):
    areas = report.get("areas") or []
    if not areas:
        raise ValueError("Bitte mindestens einen Arbeitsbereich auswählen")
    usage = report.get("usage") or {}
    for area in areas:
        area_usage = usage.get(area) or {}
        for kind in ("work", "vehicles", "equipment", "materials"):
            if area_usage.get(kind) not in {"used", "none"}:
                raise ValueError(f"Bitte für {area}/{kind} aktiv 'benutzt' oder 'nichts benutzt' auswählen")
        if area_usage.get("work") == "used":
            entries = (report.get("work") or {}).get(area) or []
            if not any(report_person_name(x) and float(x.get("hours") or 0) > 0 and x.get("description") for x in entries):
                raise ValueError("Bitte Mitarbeitende, Arbeitsbeschrieb und Stunden vollständig erfassen")
        if area_usage.get("vehicles") == "used" and not (report.get("vehicles") or {}).get(area):
            raise ValueError("Bitte ein Fahrzeug erfassen oder 'nichts benutzt' auswählen")
        if area_usage.get("equipment") == "used" and not any(x.get("item") for x in (report.get("equipment") or {}).get(area, [])):
            raise ValueError("Bitte ein Gerät erfassen oder 'nichts benutzt' auswählen")
        if area_usage.get("materials") == "used":
            material = list((report.get("products") or {}).get(area, [])) + list((report.get("consumables") or {}).get(area, []))
            disposal = (report.get("disposal") or {}).get(area) or {}
            if not any(float(x.get("amount") or 0) > 0 for x in material) and not float(disposal.get("amount") or 0):
                raise ValueError("Bitte Material erfassen oder 'nichts benutzt' auswählen")


def report_person_name(entry):
    full = str(entry.get("name") or "").strip()
    if full:
        return full
    return " ".join(part for part in (str(entry.get("firstName") or "").strip(), str(entry.get("lastName") or "").strip()) if part)


def calculate_costs(report):
    work = [item for entries in (report.get("work") or {}).values() for item in entries]
    labour = sum(float(item.get("hours") or 0) * WORK_PRICES.get(item.get("role"), 0) for item in work)
    vehicles = 0
    for entries in (report.get("vehicles") or {}).values():
        for item in entries:
            price = float(item.get("unitPrice") or VEHICLE_PRICES.get(item.get("type"), {}).get(item.get("billing"), 0))
            vehicles += float(item.get("count") or 0) * float(item.get("amount") or 0) * price
    equipment = 0
    for entries in (report.get("equipment") or {}).values():
        for item in entries:
            quantity = float(item.get("quantity") or item.get("hours") or 0)
            if not quantity and item.get("built") and item.get("removed"):
                try:
                    quantity = max(1, (datetime.fromisoformat(item["removed"]) - datetime.fromisoformat(item["built"])).days + 1)
                except ValueError:
                    quantity = 0
            equipment += quantity * float(item.get("unitPrice") or 0)
    materials = 0
    for section in ("products", "consumables"):
        for entries in (report.get(section) or {}).values():
            for item in entries:
                materials += float(item.get("amount") or 0) * float(item.get("unitPrice") or MATERIAL_PRICES.get(item.get("item"), 0))
    disposal = sum(float(item.get("amount") or 0) * (130 if item.get("unit") == "m³" else 0) for item in (report.get("disposal") or {}).values() if isinstance(item, dict))
    third_party = sum(float(item.get("price") or 0) for item in report.get("thirdParty") or [])
    total = labour + vehicles + equipment + materials + disposal + third_party
    revenue = float((report.get("costs") or {}).get("invoiceAmount") or 0)
    margin = revenue - total
    return {"labour": labour, "vehicles": vehicles, "equipment": equipment, "materials": materials + disposal,
            "thirdParty": third_party, "total": total, "revenue": revenue, "margin": margin,
            "marginPercent": (margin / revenue * 100) if revenue else 0}


def clear_and_mark(ws, choices, selected, text_cell=None):
    for label, cell in choices.items():
        ws[cell] = "x" if label == selected else None
    if text_cell and selected not in choices:
        ws[text_cell] = selected


def write_rows(ws, rows, start_row, columns, max_rows):
    for offset in range(max_rows):
        row_number = start_row + offset
        item = rows[offset] if offset < len(rows) else {}
        for key, column in columns.items():
            value = item.get(key, "")
            ws[f"{column}{row_number}"] = value if value not in (None, "") else None


def preserve_template_package(template_path, generated_bytes):
    """Keep the template's drawings, controls, validation extensions and print settings."""
    with zipfile.ZipFile(template_path) as source_zip, zipfile.ZipFile(io.BytesIO(generated_bytes)) as generated_zip:
        files = {name: source_zip.read(name) for name in source_zip.namelist()}
        replace_prefixes = ("xl/worksheets/sheet",)
        replace_names = {"xl/styles.xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/theme/theme1.xml"}
        for name in generated_zip.namelist():
            if name in replace_names or any(name.startswith(prefix) and name.endswith(".xml") for prefix in replace_prefixes):
                files[name] = generated_zip.read(name)
        source_sheet = ET.fromstring(source_zip.read("xl/worksheets/sheet1.xml"))
        generated_sheet = ET.fromstring(files["xl/worksheets/sheet1.xml"])
        preserved = {"pageSetup", "drawing", "legacyDrawing", "legacyDrawingHF", "controls", "extLst"}
        for child in list(generated_sheet):
            if child.tag.rsplit("}", 1)[-1] in preserved:
                generated_sheet.remove(child)
        for child in source_sheet:
            if child.tag.rsplit("}", 1)[-1] in preserved:
                generated_sheet.append(child)
        files["xl/worksheets/sheet1.xml"] = ET.tostring(generated_sheet, encoding="utf-8", xml_declaration=True)
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as final_zip:
            for name, content in files.items():
                final_zip.writestr(name, content)
        return output.getvalue()


def build_excel(report):
    language = report.get("language") if report.get("language") in TEMPLATE_PATHS else "de"
    workbook = load_workbook(TEMPLATE_PATHS[language])
    ws = workbook[workbook.sheetnames[0]]
    location, costs = report.get("location") or {}, report.get("costs") or {}
    ws["B1"] = report.get("projectLead")
    ws["B2"] = report.get("projectName")
    ws["B3"] = " ".join(x for x in (location.get("address"), location.get("city")) if x)
    ws["B4"] = report.get("date")
    ws["E4"] = report.get("projectNumber")
    ws["B5"] = report.get("damageType")
    ws["B6"] = f"Projektstatus: {report.get('projectStatus', 'Offen')}"

    work = report.get("work") or {}
    descriptions = []
    for key, title in (("fh", "FAH"), ("gb", "GEB"), ("tr", "Trocknung")):
        roles = sorted({str(item.get("description") or item.get("role") or "") for item in work.get(key, []) if item.get("description") or item.get("role")})
        if roles:
            descriptions.append(f"{title}: " + ", ".join(roles))
    for index, text in enumerate(descriptions[:5], start=10):
        ws[f"A{index}"] = text

    staff = []
    for area in ("fh", "gb", "tr"):
        for item in work.get(area, []):
            staff.append((area, item))
    for row_number, (area, item) in zip(range(17, 33), staff):
        ws[f"A{row_number}"] = " / ".join(x for x in (item.get("role"), item.get("description")) if x)
        name = report_person_name(item)
        if item.get("temp"):
            name += f" / {item['temp']}"
        ws[f"B{row_number}"] = name
        ws[f"C{row_number}"] = float(item.get("hours") or 0) if area == "fh" else None
        ws[f"D{row_number}"] = float(item.get("hours") or 0) if area != "fh" else None
        ws[f"E{row_number}"] = "Std"
        ws[f"F{row_number}"] = WORK_PRICES.get(item.get("role"), 0)

    vehicle_rows = {("PW", "Kilometer"): 37, ("PW", "Pauschale"): 38, ("Bus", "Kilometer"): 39, ("Bus", "Pauschale"): 40}
    for entries in (report.get("vehicles") or {}).values():
        for item in entries:
            row_number = vehicle_rows.get((item.get("type"), item.get("billing")))
            if row_number:
                current = float(ws[f"C{row_number}"].value or 0)
                ws[f"C{row_number}"] = current + float(item.get("count") or 0) * float(item.get("amount") or 0)
                ws[f"F{row_number}"] = float(item.get("unitPrice") or VEHICLE_PRICES[item["type"]][item["billing"]])

    equipment = [item for entries in (report.get("equipment") or {}).values() for item in entries]
    for row_number, item in zip(range(48, 58), equipment):
        ws[f"A{row_number}"] = item.get("item")
        ws[f"C{row_number}"] = float(item.get("quantity") or item.get("hours") or 0)
        ws[f"E{row_number}"] = item.get("unit")
        ws[f"F{row_number}"] = float(item.get("unitPrice") or 0)

    products = [item for section in ("products", "consumables") for entries in (report.get(section) or {}).values() for item in entries]
    for row_number, item in zip(range(61, 70), products):
        ws[f"A{row_number}"] = item.get("item")
        ws[f"C{row_number}"] = float(item.get("amount") or 0)
        ws[f"E{row_number}"] = "Einheit"
        ws[f"F{row_number}"] = float(item.get("unitPrice") or MATERIAL_PRICES.get(item.get("item"), 0))

    disposal = [item for item in (report.get("disposal") or {}).values() if isinstance(item, dict) and item.get("amount")]
    if disposal:
        ws["C80"] = sum(float(item.get("amount") or 0) for item in disposal if item.get("unit") == "m³")
    totals = calculate_costs(report)
    ws["G82"], ws["H82"] = totals["labour"], 0
    ws["G83"], ws["H83"] = 0, 0
    ws["G84"], ws["H84"] = totals["vehicles"], 0
    ws["G85"], ws["H85"] = 0, 0
    ws["G86"], ws["H86"] = totals["equipment"], 0
    ws["G87"], ws["H87"] = totals["materials"], 0
    ws["G88"], ws["H88"] = 0, 0
    ws["A89"] = "Fremdleistungen"
    ws["G89"], ws["H89"] = totals["thirdParty"], 0
    remarks = [
        f"Versicherung: {costs.get('insurance', '')}",
        f"Schadeninspektor (SI): {costs.get('inspectorName', '')} | {costs.get('inspectorEmail', '')}",
        f"Schadennummer: {costs.get('claimNumber', '')}",
        f"Verrechenbarer Betrag: CHF {totals['revenue']:.2f}",
        f"Gesamtkosten: CHF {totals['total']:.2f}",
        f"Marge: CHF {totals['margin']:.2f} ({totals['marginPercent']:.1f}%)",
    ]
    if report.get("officeWorker"):
        remarks.append(f"Zugewiesen an Office: {report['officeWorker'].get('name', '')} | {report['officeWorker'].get('email', '')}")
    for item in report.get("thirdParty") or []:
        remarks.append(f"Fremdleistung: {item.get('description', '')} | {item.get('company', '')} | CHF {float(item.get('price') or 0):.2f}")
    if report.get("notes"):
        remarks.append(f"Bemerkungen: {report['notes']}")
    for row_number, text in zip(range(95, 105), remarks):
        ws[f"A{row_number}"] = text
    ws.sheet_view.showGridLines = False
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return preserve_template_package(TEMPLATE_PATHS[language], output.read())


def pdf_table(rows, widths=None):
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#082A3C")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D8E2E7")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6F8")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return table


def build_pdf(report):
    output = io.BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="RapportTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=20, textColor=colors.HexColor("#082A3C"), alignment=TA_LEFT, spaceAfter=12))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12, textColor=colors.HexColor("#0D6C7F"), spaceBefore=12, spaceAfter=6))
    document = SimpleDocTemplate(output, pagesize=A4, rightMargin=16*mm, leftMargin=16*mm, topMargin=15*mm, bottomMargin=15*mm, title=safe_title(report))
    story = [Paragraph("Tagesrapport", styles["RapportTitle"])]
    location = report.get("location") or {}
    basics = [
        ["Datum", report.get("date", "")], ["Ausgefüllt von", report.get("reporterName", "")],
        ["Region", report.get("region", "")], ["Projektleiter", report.get("projectLead", "")],
        ["Schadenart", report.get("damageType", "")], ["Projektnummer", report.get("projectNumber", "")],
        ["Projektname", report.get("projectName", location.get("name", ""))], ["Projektstatus", report.get("projectStatus", "Offen")],
        ["Adresse", location.get("address", "")], ["Ort", location.get("city", "")],
        ["Bereiche", ", ".join({"fh": "Fahrhabe", "gb": "Gebäude", "tr": "Trocknung"}.get(x, x) for x in (report.get("areas") or []))],
    ]
    if report.get("officeWorker"):
        basics.append(["Zugewiesen an Office", f"{report['officeWorker'].get('name', '')} ({report['officeWorker'].get('email', '')})"])
    story.append(pdf_table([["Angabe", "Wert"]] + basics, [48*mm, 125*mm]))

    area_names = {"fh": "Fahrhabe", "gb": "Gebäude", "tr": "Trocknung"}
    for key in report.get("areas") or []:
        story.extend([Paragraph(area_names.get(key, key), styles["Section"]), Paragraph("Arbeitsstunden", styles["Heading3"])])
        work_rows = [["Funktion", "Arbeitsbeschrieb", "Personaltyp", "Nachname", "Vorname", "Stunden", "Temporärbüro"]]
        for item in (report.get("work") or {}).get(key, []):
            if report_person_name(item) or item.get("hours"):
                work_rows.append([item.get("role", ""), item.get("description", item.get("role", "")), item.get("employeeType", ""), item.get("lastName", item.get("name", "")), item.get("firstName", ""), item.get("hours", ""), item.get("temp", "")])
        if len(work_rows) > 1:
            story.append(pdf_table(work_rows, [22*mm, 32*mm, 20*mm, 26*mm, 24*mm, 13*mm, 31*mm]))
        extra = (report.get("extras") or {}).get(key) or {}
        extra_rows = [["Zuschlag / Pauschale", "Wert"]]
        if extra.get("saturday") not in ("", None, 0):
            extra_rows.append(["Samstag (Std.)", extra.get("saturday")])
        if extra.get("sunday") not in ("", None, 0):
            extra_rows.append(["Sonntag (Std.)", extra.get("sunday")])
        if extra.get("leadFlat"):
            extra_rows.append(["Projektleiter-Pauschale", extra.get("leadFlat")])
        if len(extra_rows) > 1:
            story.append(Spacer(1, 2*mm))
            story.append(pdf_table(extra_rows, [70*mm, 98*mm]))

        for section_key, title, columns in [
            ("vehicles", "Fahrzeuge", [("type", "Fahrzeug"), ("count", "Anzahl"), ("billing", "Abrechnung"), ("amount", "KM / Stufen"), ("unitPrice", "Fixpreis CHF")]),
            ("equipment", "Geräte", [("item", "Gerät"), ("quantity", "Anzahl"), ("unit", "Einheit"), ("unitPrice", "Preis CHF"), ("built", "Aufgebaut"), ("removed", "Abgebaut")]),
            ("products", "Sanierungsprodukte / Spezialgeräte", [("item", "Position"), ("amount", "Menge"), ("unitPrice", "Preis CHF")]),
            ("consumables", "Verbrauchsmaterial", [("item", "Position"), ("amount", "Menge"), ("unitPrice", "Preis CHF")]),
        ]:
            items = [item for item in (report.get(section_key) or {}).get(key, []) if any(v not in ("", None, 0) for v in item.values())]
            if items:
                story.append(Paragraph(title, styles["Heading3"]))
                story.append(pdf_table([[label for _, label in columns]] + [[item.get(field, "") for field, _ in columns] for item in items]))
        disposal = (report.get("disposal") or {}).get(key) or {}
        if isinstance(disposal, dict) and disposal.get("method"):
            story.append(Paragraph("Entsorgung", styles["Heading3"]))
            story.append(pdf_table([["Entsorgungsweg", "Menge", "Einheit"], [disposal.get("method", ""), disposal.get("amount", ""), disposal.get("unit", "")]]))
        story.append(Spacer(1, 4*mm))

    costs, totals = report.get("costs") or {}, calculate_costs(report)
    story.extend([PageBreak(), Paragraph("Kostenübersicht Projektleitung", styles["RapportTitle"])])
    insurance_rows = [
        ["Versicherung", costs.get("insurance", "")], ["Schadeninspektor (SI)", costs.get("inspectorName", "")],
        ["E-Mail SI", costs.get("inspectorEmail", "")], ["Schadennummer", costs.get("claimNumber", "")],
    ]
    story.append(pdf_table([["Angabe", "Wert"]] + insurance_rows, [58*mm, 115*mm]))
    third_party = report.get("thirdParty") or []
    if third_party:
        story.extend([Paragraph("Fremdleistungen", styles["Section"]), pdf_table([["Leistung / Beschreibung", "Firma", "Preis CHF"]] + [[item.get("description", ""), item.get("company", ""), f"{float(item.get('price') or 0):.2f}"] for item in third_party], [75*mm, 62*mm, 35*mm])])
    story.append(Paragraph("Kalkulation", styles["Section"]))
    cost_rows = [
        ["Personalkosten", f"CHF {totals['labour']:.2f}"], ["Fahrzeugkosten", f"CHF {totals['vehicles']:.2f}"],
        ["Gerätekosten", f"CHF {totals['equipment']:.2f}"], ["Material & Entsorgung", f"CHF {totals['materials']:.2f}"],
        ["Fremdleistungen", f"CHF {totals['thirdParty']:.2f}"], ["Gesamtkosten", f"CHF {totals['total']:.2f}"],
        ["Verrechenbarer Betrag", f"CHF {totals['revenue']:.2f}"], ["Marge", f"CHF {totals['margin']:.2f} ({totals['marginPercent']:.1f}%)"],
    ]
    story.append(pdf_table([["Kostenart", "Betrag"]] + cost_rows, [105*mm, 68*mm]))

    notes = str(report.get("notes") or "").strip()
    if notes:
        story.extend([Paragraph("Bemerkungen", styles["Section"]), Paragraph(xml_escape(notes).replace("\n", "<br/>"), styles["BodyText"])])
    story.extend([Spacer(1, 8*mm), Paragraph(f"Status: {report.get('status', 'Entwurf')}", styles["BodyText"])])
    document.build(story)
    return output.getvalue()


def save_pdf_to_folder(report):
    today = datetime.now(ZoneInfo("Europe/Zurich")).date().isoformat()
    relative = Path("project_leaders") / ascii_slug(report.get("projectLead")) / today / f"rapport-{report['id']}.pdf"
    destination = DATA_DIR / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(build_pdf(report))
    return str(relative)


def save_pdf_to_frontoffice(report):
    today = datetime.now(ZoneInfo("Europe/Zurich")).date().isoformat()
    office_name = (report.get("officeWorker") or {}).get("name") or "nicht-zugewiesen"
    relative = Path("frontoffice") / ascii_slug(office_name) / today / f"rapport-{report['id']}.pdf"
    destination = DATA_DIR / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(build_pdf(report))
    return str(relative)


def public_base_url(environ):
    configured = os.environ.get("BASE_URL", "").strip().rstrip("/")
    if configured:
        return configured
    scheme = (environ.get("HTTP_X_FORWARDED_PROTO") or environ.get("wsgi.url_scheme") or "https").split(",")[0].strip()
    host = (environ.get("HTTP_X_FORWARDED_HOST") or environ.get("HTTP_HOST") or "localhost").split(",")[0].strip()
    if scheme not in {"http", "https"} or not re.fullmatch(r"[A-Za-z0-9.:-]+", host):
        raise ValueError("BASE_URL ist nicht korrekt konfiguriert")
    return f"{scheme}://{host}"


def send_project_leader_email(report, environ):
    host = os.environ.get("SMTP_HOST", "").strip()
    username = os.environ.get("SMTP_USERNAME", "").strip()
    password = os.environ.get("SMTP_PASSWORD", "")
    sender = os.environ.get("SMTP_FROM", username).strip()
    if not all([host, username, password, sender]):
        return False, "SMTP ist nicht vollständig konfiguriert"
    port = int(os.environ.get("SMTP_PORT", "587"))
    recipient = leader_email(report.get("projectLead", ""))
    reporter = report.get("reporterName") or "Eine Person"
    link = f"{public_base_url(environ)}/?portal=1&report={report['id']}"
    message = EmailMessage()
    message["Subject"] = f"Neuer Tagesrapport: {safe_title(report)}"
    message["From"] = sender
    message["To"] = recipient
    message.set_content(f"{reporter} hat einen Rapport ausgefüllt.\n\nRapport öffnen und prüfen:\n{link}\n")
    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=20) as smtp:
                smtp.login(username, password)
                smtp.send_message(message)
        else:
            with smtplib.SMTP(host, port, timeout=20) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(username, password)
                smtp.send_message(message)
        return True, recipient
    except Exception as error:
        print(f"Email error for report {report.get('id')}: {error}", flush=True)
        return False, "E-Mail konnte nicht gesendet werden"


def serve_static(start_response, path):
    mapping = {
        "/": (STATIC_DIR / "index.html", "text/html; charset=utf-8"),
        "/styles.css": (STATIC_DIR / "styles.css", "text/css; charset=utf-8"),
        "/app.js": (STATIC_DIR / "app.js", "application/javascript; charset=utf-8"),
    }
    item = mapping.get(path)
    if not item or not item[0].exists():
        return text_response(start_response, "Nicht gefunden", "text/plain; charset=utf-8", "404 Not Found")
    return text_response(start_response, item[0].read_bytes(), item[1])


def application(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO", "/")

    if path == "/health":
        return json_response(start_response, {"status": "ok"})
    if path == "/api/config":
        return json_response(start_response, {
            "authRequired": bool(os.environ.get("APP_PASSWORD", "").strip()),
            "projectLeaders": [{"name": name, "email": leader_email(name)} for name in PROJECT_LEADERS],
            "languages": ["de", "fr", "it"],
        })

    try:
        if path == "/api/auth/login" and method == "POST":
            data = read_json(environ)
            name = str(data.get("name") or "").strip()
            password = str(data.get("password") or "")
            with db_connection() as connection:
                leader = connection.execute("SELECT * FROM project_leaders WHERE name = ? COLLATE NOCASE", (name,)).fetchone()
                if not leader or not leader["password_hash"] or not leader["password_salt"]:
                    return json_response(start_response, {"error": "Login nicht eingerichtet oder ungültig"}, "401 Unauthorized")
                digest, _ = hash_password(password, leader["password_salt"])
                if not hmac.compare_digest(digest, leader["password_hash"]):
                    return json_response(start_response, {"error": "Login nicht eingerichtet oder ungültig"}, "401 Unauthorized")
                token = secrets.token_urlsafe(36)
                expires = datetime.now(timezone.utc) + timedelta(hours=12)
                connection.execute("DELETE FROM login_sessions WHERE expires_at <= ?", (utc_now(),))
                connection.execute(
                    "INSERT INTO login_sessions (token_hash, project_leader_id, expires_at) VALUES (?, ?, ?)",
                    (hashlib.sha256(token.encode("utf-8")).hexdigest(), leader["id"], expires.replace(microsecond=0).isoformat()),
                )
                connection.commit()
            return json_response(start_response, {"token": token, "projectLead": leader["name"], "expiresAt": expires.isoformat()})

        context = auth_context(environ)
        if path.startswith("/api/") and not context:
            return json_response(start_response, {"error": "Zugriffsschlüssel oder Login ungültig"}, "401 Unauthorized")

        if path == "/api/project-leaders" and method == "GET":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            with db_connection() as connection:
                rows = connection.execute("SELECT id, name, email, password_hash FROM project_leaders ORDER BY name").fetchall()
            return json_response(start_response, {"projectLeaders": [{"id": row["id"], "name": row["name"], "email": row["email"], "loginConfigured": bool(row["password_hash"])} for row in rows]})

        if path == "/api/project-leaders/password" and method == "POST":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            data = read_json(environ)
            name = str(data.get("name") or "").strip()
            password = str(data.get("password") or "")
            if len(password) < 10:
                raise ValueError("Das Passwort muss mindestens 10 Zeichen lang sein")
            digest, salt = hash_password(password)
            with db_connection() as connection:
                result = connection.execute("UPDATE project_leaders SET password_hash = ?, password_salt = ?, updated_at = ? WHERE name = ? COLLATE NOCASE", (digest, salt, utc_now(), name))
                connection.execute("DELETE FROM login_sessions WHERE project_leader_id IN (SELECT id FROM project_leaders WHERE name = ? COLLATE NOCASE)", (name,))
                connection.commit()
            if not result.rowcount:
                raise ValueError("Projektleiter nicht gefunden")
            return json_response(start_response, {"updated": True, "name": name})

        if path == "/api/reports" and method == "GET":
            with db_connection() as connection:
                rows = connection.execute("SELECT id, title, status, created_at, updated_at, payload FROM reports ORDER BY updated_at DESC LIMIT 250").fetchall()
            reports = []
            for row in rows:
                item = dict(row)
                payload = json.loads(item.pop("payload"))
                if context["kind"] == "leader" and payload.get("projectLead") != context["name"]:
                    continue
                item["projectLead"] = payload.get("projectLead", "Ohne Zuteilung")
                item["date"] = payload.get("date", "")
                item["location"] = (payload.get("location") or {}).get("name", "")
                item["projectNumber"] = payload.get("projectNumber", "")
                item["projectName"] = payload.get("projectName") or item["location"]
                item["projectStatus"] = payload.get("projectStatus", "Offen")
                item["reporterName"] = payload.get("reporterName", "")
                item["officeWorker"] = (payload.get("officeWorker") or {}).get("name", "")
                item["folderDate"] = payload.get("_folderDate") or payload.get("date", "")
                item["hasPdf"] = bool(payload.get("_pdfPath"))
                reports.append(item)
            return json_response(start_response, {"reports": reports, "scope": context["kind"], "projectLead": context.get("name")})

        if path == "/api/reports" and method == "POST":
            data = normalize_report(read_json(environ))
            report_id = str(data.get("id") or uuid.uuid4())
            if not re.fullmatch(r"[0-9a-fA-F-]{36}", report_id):
                report_id = str(uuid.uuid4())
            now = utc_now()
            data["id"] = report_id
            title = safe_title(data)
            status = str(data.get("status") or "Entwurf")[:40]
            with db_connection() as connection:
                existing = connection.execute("SELECT created_at, payload FROM reports WHERE id = ?", (report_id,)).fetchone()
                if existing:
                    previous = json.loads(existing["payload"])
                    for metadata_key in ("_pdfPath", "_frontofficePath", "_folderDate", "_submittedAt"):
                        if previous.get(metadata_key) and not data.get(metadata_key):
                            data[metadata_key] = previous[metadata_key]
                if context["kind"] == "leader":
                    if not existing:
                        return json_response(start_response, {"error": "Projektleiter können nur zugeteilte Rapporte bearbeiten"}, "403 Forbidden")
                    if previous.get("projectLead") != context["name"]:
                        return json_response(start_response, {"error": "Keine Berechtigung für diesen Rapport"}, "403 Forbidden")
                    data["projectLead"] = previous.get("projectLead")
                    data["_pdfPath"] = previous.get("_pdfPath")
                    data["_frontofficePath"] = previous.get("_frontofficePath")
                    data["_folderDate"] = previous.get("_folderDate")
                    status = previous.get("status", "Zur Kontrolle")
                    data["status"] = status
                created_at = existing["created_at"] if existing else now
                connection.execute(
                    "INSERT OR REPLACE INTO reports (id, title, status, created_at, updated_at, payload) VALUES (?, ?, ?, ?, ?, ?)",
                    (report_id, title, status, created_at, now, json.dumps(data, ensure_ascii=False)),
                )
                connection.commit()
            if data.get("_pdfPath"):
                pdf_path = DATA_DIR / data["_pdfPath"]
                pdf_path.parent.mkdir(parents=True, exist_ok=True)
                pdf_path.write_bytes(build_pdf(data))
            return json_response(start_response, {"id": report_id, "title": title, "status": status}, "201 Created")

        if path == "/api/employees" and method == "GET":
            with db_connection() as connection:
                rows = connection.execute("SELECT id, name FROM employees WHERE active = 1 ORDER BY name COLLATE NOCASE").fetchall()
                report_rows = connection.execute("SELECT payload FROM reports").fetchall()
            people = {row["name"].casefold(): {"id": row["id"], "name": row["name"], "source": "saved"} for row in rows}
            for report_row in report_rows:
                report = json.loads(report_row["payload"])
                for entries in (report.get("work") or {}).values():
                    for entry in entries:
                        name = report_person_name(entry)
                        if name and name.casefold() not in people:
                            people[name.casefold()] = {"id": None, "name": name, "source": "report"}
            return json_response(start_response, {"employees": sorted(people.values(), key=lambda item: item["name"].casefold())})

        if path == "/api/office-workers" and method == "GET":
            with db_connection() as connection:
                rows = connection.execute("SELECT id, name, email FROM office_workers WHERE active = 1 ORDER BY name COLLATE NOCASE").fetchall()
            return json_response(start_response, {"officeWorkers": [dict(row) for row in rows]})

        if path == "/api/office-workers" and method == "POST":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            data = read_json(environ)
            name, email = str(data.get("name") or "").strip(), str(data.get("email") or "").strip().lower()
            if len(name) < 2 or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                raise ValueError("Bitte Name und gültige E-Mail eingeben")
            worker_id = str(uuid.uuid4())
            with db_connection() as connection:
                connection.execute("INSERT INTO office_workers (id, name, email, created_at) VALUES (?, ?, ?, ?)", (worker_id, name, email, utc_now()))
                connection.commit()
            return json_response(start_response, {"officeWorker": {"id": worker_id, "name": name, "email": email}}, "201 Created")

        office_match = re.fullmatch(r"/api/office-workers/([0-9a-fA-F-]{36})", path)
        if office_match and method == "DELETE":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            with db_connection() as connection:
                connection.execute("UPDATE office_workers SET active = 0 WHERE id = ?", (office_match.group(1),))
                connection.commit()
            return json_response(start_response, {"deleted": True})

        if path == "/api/employees" and method == "POST":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            data = read_json(environ)
            name = str(data.get("name") or "").strip()
            if len(name) < 2:
                raise ValueError("Bitte einen gültigen Namen eingeben")
            employee_id = str(uuid.uuid4())
            with db_connection() as connection:
                try:
                    connection.execute("INSERT INTO employees (id, name, created_at) VALUES (?, ?, ?)", (employee_id, name, utc_now()))
                    connection.commit()
                except sqlite3.IntegrityError:
                    raise ValueError("Dieser Mitarbeiter ist bereits gespeichert")
            return json_response(start_response, {"employee": {"id": employee_id, "name": name}}, "201 Created")

        employee_match = re.fullmatch(r"/api/employees/([0-9a-fA-F-]{36})", path)
        if employee_match and method == "DELETE":
            if context["kind"] != "admin":
                return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
            with db_connection() as connection:
                connection.execute("UPDATE employees SET active = 0 WHERE id = ?", (employee_match.group(1),))
                connection.commit()
            return json_response(start_response, {"deleted": True})

        if path == "/api/hour-summary" and method == "GET":
            query = parse_qs(environ.get("QUERY_STRING", ""))
            wanted = {key: (query.get(key) or [""])[0] for key in ("from", "to", "region", "employee", "project", "lead", "area", "projectStatus", "reportStatus")}
            details, all_details = [], []
            with db_connection() as connection:
                rows = connection.execute("SELECT status, payload FROM reports").fetchall()
            for row in rows:
                report = json.loads(row["payload"])
                if context["kind"] == "leader" and report.get("projectLead") != context["name"]:
                    continue
                for area, entries in (report.get("work") or {}).items():
                    for entry in entries:
                        name = report_person_name(entry)
                        hours = float(entry.get("hours") or 0)
                        if not name or not hours:
                            continue
                        item = {"date": report.get("date", ""), "region": report.get("region", ""), "name": name, "area": area, "hours": round(hours, 2), "description": entry.get("description") or entry.get("role", ""), "projectNumber": report.get("projectNumber", ""), "projectName": report.get("projectName", ""), "projectLead": report.get("projectLead", ""), "projectStatus": report.get("projectStatus", ""), "reportStatus": row["status"]}
                        all_details.append(item)
                        project_key = f"{item['projectNumber']} · {item['projectName']}"
                        if wanted["from"] and item["date"] < wanted["from"] or wanted["to"] and item["date"] > wanted["to"]: continue
                        if wanted["region"] and item["region"] != wanted["region"] or wanted["employee"] and item["name"] != wanted["employee"]: continue
                        if wanted["project"] and project_key != wanted["project"] or wanted["lead"] and item["projectLead"] != wanted["lead"]: continue
                        if wanted["area"] and item["area"] != wanted["area"] or wanted["projectStatus"] and item["projectStatus"] != wanted["projectStatus"] or wanted["reportStatus"] and item["reportStatus"] != wanted["reportStatus"]: continue
                        details.append(item)
            totals = {}
            for item in details:
                current = totals.setdefault((item["region"], item["name"]), {"name": item["name"], "fh": 0, "gb": 0, "tr": 0, "total": 0})
                current[item["area"]] = round(current[item["area"]] + item["hours"], 2)
                current["total"] = round(current["total"] + item["hours"], 2)
            regions = {}
            for (region, _), item in totals.items(): regions.setdefault(region or "Ohne Region", []).append(item)
            pivot = [{"region": region, "rows": sorted(items, key=lambda x: x["name"].casefold())} for region, items in sorted(regions.items())]
            filters = {"regions": sorted({x["region"] for x in all_details}), "employees": sorted({x["name"] for x in all_details}), "projects": sorted({f"{x['projectNumber']} · {x['projectName']}" for x in all_details}), "leads": sorted({x["projectLead"] for x in all_details}), "projectStatuses": sorted({x["projectStatus"] for x in all_details}), "reportStatuses": sorted({x["reportStatus"] for x in all_details})}
            details.sort(key=lambda x: (x["date"], x["region"], x["name"]), reverse=True)
            return json_response(start_response, {"pivot": pivot, "details": details, "filterOptions": filters})

        match = re.fullmatch(r"/api/reports/([0-9a-fA-F-]{36})(/(?:excel|pdf))?", path)
        if match:
            report_id, file_suffix = match.group(1), match.group(2)
            with db_connection() as connection:
                row = connection.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
            if not row:
                return json_response(start_response, {"error": "Rapport nicht gefunden"}, "404 Not Found")
            report = json.loads(row["payload"])
            if context["kind"] == "leader" and report.get("projectLead") != context["name"]:
                return json_response(start_response, {"error": "Keine Berechtigung für diesen Rapport"}, "403 Forbidden")
            if method == "GET" and file_suffix == "/excel":
                binary = build_excel(report)
                filename = re.sub(r"[^A-Za-z0-9._-]+", "_", row["title"]) + ".xlsx"
                start_response("200 OK", [
                    ("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    ("Content-Disposition", f'attachment; filename="{filename}"'),
                    ("Content-Length", str(len(binary))),
                ])
                return [binary]
            if method == "GET" and file_suffix == "/pdf":
                relative = report.get("_pdfPath")
                binary = (DATA_DIR / relative).read_bytes() if relative and (DATA_DIR / relative).exists() else build_pdf(report)
                filename = re.sub(r"[^A-Za-z0-9._-]+", "_", row["title"]) + ".pdf"
                start_response("200 OK", [("Content-Type", "application/pdf"), ("Content-Disposition", f'attachment; filename="{filename}"'), ("Content-Length", str(len(binary)))])
                return [binary]
            if method == "GET" and not file_suffix:
                return json_response(start_response, {"report": report})
            if method == "DELETE" and not file_suffix:
                if context["kind"] != "admin":
                    return json_response(start_response, {"error": "Keine Berechtigung"}, "403 Forbidden")
                with db_connection() as connection:
                    connection.execute("DELETE FROM reports WHERE id = ?", (report_id,))
                    connection.commit()
                return json_response(start_response, {"deleted": True})

        workflow_match = re.fullmatch(r"/api/reports/([0-9a-fA-F-]{36})/workflow", path)
        if workflow_match and method == "POST":
            data = read_json(environ)
            requested = str(data.get("status") or "")
            allowed = {"Entwurf", "Zur Kontrolle", "Frontoffice", "Erledigt"}
            if requested not in allowed:
                raise ValueError("Ungültiger Workflow-Status")
            with db_connection() as connection:
                row = connection.execute("SELECT payload FROM reports WHERE id = ?", (workflow_match.group(1),)).fetchone()
                if not row:
                    return json_response(start_response, {"error": "Rapport nicht gefunden"}, "404 Not Found")
                payload = json.loads(row["payload"])
                if context["kind"] == "leader":
                    if payload.get("projectLead") != context["name"] or requested not in {"Zur Kontrolle", "Frontoffice"}:
                        return json_response(start_response, {"error": "Keine Berechtigung für diesen Status"}, "403 Forbidden")
                if requested == "Zur Kontrolle":
                    if not str((payload.get("costs") or {}).get("insurance") or "").strip():
                        raise ValueError("Vor dem Übermitteln muss die Versicherung angegeben werden")
                    validate_submission(payload)
                    folder_date = datetime.now(ZoneInfo("Europe/Zurich")).date().isoformat()
                    payload["_folderDate"] = folder_date
                    payload["_submittedAt"] = utc_now()
                payload["status"] = requested
                if requested == "Zur Kontrolle":
                    payload["_pdfPath"] = save_pdf_to_folder(payload)
                if requested == "Frontoffice":
                    worker_id = str(data.get("officeWorkerId") or "")
                    worker = connection.execute("SELECT id, name, email FROM office_workers WHERE id = ? AND active = 1", (worker_id,)).fetchone()
                    if not worker:
                        raise ValueError("Bitte eine aktive Office-Person auswählen")
                    payload["officeWorker"] = {**dict(worker), "assignedAt": utc_now()}
                    payload["_frontofficePath"] = save_pdf_to_frontoffice(payload)
                connection.execute("UPDATE reports SET status = ?, payload = ?, updated_at = ? WHERE id = ?", (requested, json.dumps(payload, ensure_ascii=False), utc_now(), workflow_match.group(1)))
                connection.commit()
            mail_sent, mail_detail = (False, None)
            if requested == "Zur Kontrolle":
                mail_sent, mail_detail = send_project_leader_email(payload, environ)
            return json_response(start_response, {"id": workflow_match.group(1), "status": requested, "pdfCreated": bool(payload.get("_pdfPath")), "mailSent": mail_sent, "mailDetail": mail_detail})

        if path.startswith("/api/"):
            return json_response(start_response, {"error": "Nicht gefunden"}, "404 Not Found")
        return serve_static(start_response, path)
    except (ValueError, json.JSONDecodeError) as error:
        return json_response(start_response, {"error": str(error)}, "400 Bad Request")
    except Exception as error:
        print(f"Request error: {error}", flush=True)
        return json_response(start_response, {"error": "Interner Fehler"}, "500 Internal Server Error")


if __name__ == "__main__":
    from wsgiref.simple_server import make_server

    port = int(os.environ.get("PORT", "8080"))
    print(f"Tagesrapport läuft auf http://localhost:{port}")
    make_server("0.0.0.0", port, application).serve_forever()
