import os
import re
from datetime import datetime, timedelta, timezone
import xml.etree.ElementTree as ET
import pandas as pd
import requests

# -------------------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------------------
XML_OUTPUT_FILE = "CNLA_EPG.xml"
CHANNEL_ID = "CNLA_PAN.co"
CHANNEL_NAME = "Cartoon Network Panregional"
RETENTION_DAYS = 15
COT = timezone(timedelta(hours=-5))
DAILY_SHOWS_URL = "https://epg.tapkit.warnermedia.com/api/daily/shows?feedId=CNLA_PAN&format=xls"

def download_panregional_xls():
    token = os.environ.get("TAPKIT_TOKEN")
    if not token:
        raise Exception("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyX25hbWUiOiJjZW9AYnNtYWdlbmN5LmNvbS5jbyIsInNjb3BlIjpbImFueSJdLCJsYXN0X25hbWUiOiJNZXJjYWRvIiwic3RhdGUiOjMsImlkIjoyMjIxNSwiZXhwIjoxNzg3MjY3NzI2LCJmaXJzdF9uYW1lIjoiSmVhbiBQaGlsaXBwIiwiYXV0aG9yaXRpZXMiOlsiUk9MRV9BRE1JTl9DT05URU5UIiwiUk9MRV9BTkNJTkUiLCJST0xFX0xPQ0FMQlJFQUtTIiwiUk9MRV9BRE1JTl9VU0VSUyIsIlJPTEVfVVNFUiIsIlJPTEVfQURNSU5fUFJFU1MiLCJST0xFX01JQ1JPU0VSVklDRVMiXSwianRpIjoiYjMyYmQ2NGQtNDRmYy00ZWZjLWIzODgtOGUyMWU5MzEzYWY3IiwiY2xpZW50X2lkIjoibXljbGllbnQifQ.ZCqK0sV6NTE0en5bP1BcoxZNyeg2Fsd1q9ZqcubpKhM")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "es",
        "Authorization": f"Bearer {token.strip()}",
        "Referer": "https://epg.tapkit.warnermedia.com/epg/networks/2"
    }

    res = requests.get(DAILY_SHOWS_URL, headers=headers, timeout=60)
    res.raise_for_status()

    xls_path = "CNLA_PAN_latest.xls"
    with open(xls_path, "wb") as f:
        f.write(res.content)

    print("[OK] XLS descargado exitosamente.")
    return xls_path

def sanitize_and_parse_xml(file_path):
    if not os.path.exists(file_path):
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network Panregional"})

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        cleaned_lines = []
        for line in content.splitlines():
            stripped = line.strip()
            if stripped in ["JUEVES", "VIERNES", "SÁBADO", "SABADO", "DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "MIERCOLES"] or ("PANREGIONAL" in stripped and not stripped.startswith("<")):
                continue
            cleaned_lines.append(line)

        cleaned_content = "\n".join(cleaned_lines)
        return ET.fromstring(cleaned_content)
    except Exception:
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network Panregional"})

def parse_xmltv_date(date_str):
    if not date_str:
        return None
    clean_str = re.sub(r"[^\d\s\+\-]", "", str(date_str).strip())
    match = re.match(r"^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", clean_str)
    if not match:
        return None

    year, month, day, hour, minute, second = map(int, match.groups())
    hour, minute, second = min(hour, 23), min(minute, 59), min(second, 59)
    return datetime(year, month, day, hour, minute, second, tzinfo=COT)

def format_xmltv_date(dt):
    return dt.strftime("%Y%m%d%H%M%S -0500")

def load_excel_schedule(file_path):
    try:
        df = pd.read_excel(file_path, header=1, sheet_name=0)
    except Exception:
        df = pd.read_html(file_path, header=1)[0]
    return df

def update_epg_xml(xls_path):
    df = load_excel_schedule(xls_path)
    root = sanitize_and_parse_xml(XML_OUTPUT_FILE)

    # 1. Mantener canal CNLA_PAN.co
    for ch in list(root.findall("channel")):
        if ch.attrib.get("id") != CHANNEL_ID:
            root.remove(ch)

    if not any(ch.attrib.get("id") == CHANNEL_ID for ch in root.findall("channel")):
        ch_node = ET.SubElement(root, "channel", {"id": CHANNEL_ID})
        disp = ET.SubElement(ch_node, "display-name")
        disp.text = CHANNEL_NAME

    # 2. Identificar columnas
    cols = {str(c).strip(): c for c in df.columns}
    col_date = cols.get("Schedule Date", df.columns[0])
    col_time = cols.get("Title Start Time", df.columns[1])
    col_title = cols.get("Title Name", df.columns[2])
    col_ep = cols.get("Episode Name", df.columns[3] if len(df.columns) > 3 else None)
    col_desc = cols.get("Title Synopsis", df.columns[4] if len(df.columns) > 4 else None)
    col_ep_desc = cols.get("Episode Synopsis", None)

    # 3. Detectar fecha base del archivo
    first_date_raw = str(df.iloc[0].get(col_date, "")).strip()
    match_init = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", first_date_raw)
    if not match_init:
        raise Exception("No se pudo detectar la fecha inicial del archivo.")

    d0, m0, y0 = map(int, match_init.groups())
    cycle_start = datetime(y0, m0, d0, 6, 0, 0, tzinfo=COT)
    cycle_end = cycle_start + timedelta(days=1)  # 06:00 AM del día siguiente

    raw_events = []
    total_rows = len(df)

    for i in range(total_rows):
        row = df.iloc[i]
        date_raw = str(row.get(col_date, "")).strip()
        time_raw = str(row.get(col_time, "")).strip()

        date_match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", date_raw)
        time_match = re.search(r"(\d{1,2}):(\d{2})", time_raw)

        if not date_match or not time_match:
            continue

        d, m, y = map(int, date_match.groups())
        hh, mm = map(int, time_match.groups())

        # Si la hora es de madrugada (00:00 a 05:59) pero la fecha dice el mismo día base,
        # en la grilla televisiva corresponde a la madrugada del día siguiente
        event_dt = datetime(y, m, d, hh, mm, 0, tzinfo=COT)
        if (d == d0 and m == m0 and y == y0) and hh < 6:
            event_dt += timedelta(days=1)

        # Si excede el ciclo de 24h (a partir de las 06:00 AM siguientes), finaliza el día
        if event_dt >= cycle_end:
            break

        title_val = str(row.get(col_title, "")).strip()
        ep_val = str(row.get(col_ep, "")).strip() if col_ep and pd.notna(row.get(col_ep)) else ""
        
        desc_val = ""
        if col_ep_desc and pd.notna(row.get(col_ep_desc)) and str(row.get(col_ep_desc)).strip():
            desc_val = str(row.get(col_ep_desc)).strip()
        elif col_desc and pd.notna(row.get(col_desc)) and str(row.get(col_desc)).strip():
            desc_val = str(row.get(col_desc)).strip()

        raw_events.append({
            "start": event_dt,
            "title": title_val,
            "sub_title": ep_val if ep_val.lower() != title_val.lower() else "",
            "desc": desc_val
        })

    # Ordenar eventos cronológicamente (06:00 AM -> 05:59 AM del día posterior)
    raw_events = sorted(raw_events, key=lambda x: x["start"])

    # 4. Construir bloques <programme> con inicio y fin continuos
    new_programmes = []
    total_events = len(raw_events)

    for i in range(total_events):
        ev = raw_events[i]
        start_dt = ev["start"]

        if i + 1 < total_events:
            stop_dt = raw_events[i + 1]["start"]
        else:
            stop_dt = cycle_end

        prog = ET.Element("programme", {
            "start": format_xmltv_date(start_dt),
            "stop": format_xmltv_date(stop_dt),
            "channel": CHANNEL_ID
        })

        title = ET.SubElement(prog, "title", {"lang": "es"})
        title.text = ev["title"]

        if ev["sub_title"]:
            sub_title = ET.SubElement(prog, "sub-title", {"lang": "es"})
            sub_title.text = ev["sub_title"]

        if ev["desc"]:
            desc = ET.SubElement(prog, "desc", {"lang": "es"})
            desc.text = ev["desc"]

        new_programmes.append(prog)

    # 5. Fusionar programas nuevos evitando duplicados
    existing_starts = {p.attrib.get("start") for p in root.findall("programme") if p.attrib.get("channel") == CHANNEL_ID}
    for np in new_programmes:
        if np.attrib.get("start") not in existing_starts:
            root.append(np)

    # 6. Purgar eventos de más de 15 días
    cutoff_date = datetime.now(COT) - timedelta(days=RETENTION_DAYS)
    for p in list(root.findall("programme")):
        if p.attrib.get("channel") != CHANNEL_ID:
            root.remove(p)
            continue
        stop_dt = parse_xmltv_date(p.attrib.get("stop", ""))
        if stop_dt and stop_dt < cutoff_date:
            root.remove(p)

    # 7. Ordenar todo el XML cronológicamente
    sorted_progs = sorted(
        root.findall("programme"),
        key=lambda x: parse_xmltv_date(x.attrib.get("start", "")) or datetime.min.replace(tzinfo=COT)
    )

    for p in list(root.findall("programme")):
        root.remove(p)
    for p in sorted_progs:
        root.append(p)

    # 8. Guardar XML
    ET.indent(root, space="  ", level=0)
    ET.ElementTree(root).write(XML_OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    print(f"[OK] Ciclo de 24h ordenado correctamente ({len(new_programmes)} eventos del día). Total acumulado: {len(sorted_progs)}")

if __name__ == "__main__":
    xls = download_panregional_xls()
    update_epg_xml(xls)
