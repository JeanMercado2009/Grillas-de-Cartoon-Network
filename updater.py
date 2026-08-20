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

OAUTH_TOKEN_URL = "https://epg.tapkit.warnermedia.com/oauth/token"
DAILY_SHOWS_URL = "https://epg.tapkit.warnermedia.com/api/daily/shows?feedId=CNLA_PAN&format=xls"

def download_panregional_xls():
    email = os.environ.get("TAPKIT_EMAIL")
    password = os.environ.get("TAPKIT_PASSWORD")
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "es"
    })

    # 1. Autenticación OAuth2 (Cliente: myclient)
    auth_payload = {
        "grant_type": "password",
        "username": email,
        "password": password
    }
    
    auth_res = session.post(
        OAUTH_TOKEN_URL,
        data=auth_payload,
        auth=("myclient", "")
    )
    
    token = None
    if auth_res.status_code == 200:
        token = auth_res.json().get("access_token")
    else:
        # Fallback a login directo si el formato varía
        login_res = session.post(
            "https://epg.tapkit.warnermedia.com/api/auth/login",
            json={"email": email, "password": password}
        )
        if login_res.status_code == 200:
            token = login_res.json().get("accessToken") or login_res.json().get("token")

    if not token:
        raise Exception(f"Fallo en la autenticacion. Codigo de estado: {auth_res.status_code}. Respuesta: {auth_res.text}")

    # 2. Descarga del archivo XLS con el Bearer Token
    download_headers = {
        "Authorization": f"Bearer {token}",
        "Referer": "https://epg.tapkit.warnermedia.com/epg/networks/2"
    }

    xls_res = session.get(DAILY_SHOWS_URL, headers=download_headers)
    xls_res.raise_for_status()

    xls_path = "CNLA_PAN_latest.xls"
    with open(xls_path, "wb") as f:
        f.write(xls_res.content)

    print("[OK] Grilla diaria XLS descargada exitosamente vía API.")
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
    clean_str = re.sub(r"[^\d\s\+\-]", "", date_str.strip())
    match = re.match(r"^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", clean_str)
    if not match:
        return None

    year, month, day, hour, minute, second = map(int, match.groups())
    hour, minute, second = min(hour, 23), min(minute, 59), min(second, 59)
    return datetime(year, month, day, hour, minute, second, tzinfo=COT)

def format_xmltv_date(dt):
    return dt.strftime("%Y%m%d%H%M%S -0500")

def update_epg_xml(xls_path):
    df = pd.read_excel(xls_path)
    root = sanitize_and_parse_xml(XML_OUTPUT_FILE)

    # Mantener únicamente el canal CNLA_PAN.co
    for ch in list(root.findall("channel")):
        if ch.attrib.get("id") != CHANNEL_ID:
            root.remove(ch)

    if not any(ch.attrib.get("id") == CHANNEL_ID for ch in root.findall("channel")):
        ch_node = ET.SubElement(root, "channel", {"id": CHANNEL_ID})
        disp = ET.SubElement(ch_node, "display-name")
        disp.text = CHANNEL_NAME

    new_programmes = []
    total_rows = len(df)

    for i in range(total_rows):
        row = df.iloc[i]
        date_raw = str(row.get("Schedule Date", "")).strip()
        time_raw = str(row.get("Title Start Time", "")).strip()

        date_match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", date_raw)
        time_match = re.search(r"(\d{1,2}):(\d{2})", time_raw)

        if not date_match or not time_match:
            continue

        d, m, y = map(int, date_match.groups())
        hh, mm = map(int, time_match.groups())
        start_dt = datetime(y, m, d, hh, mm, 0, tzinfo=COT)

        stop_dt = None
        if i + 1 < total_rows:
            next_row = df.iloc[i + 1]
            next_d_match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", str(next_row.get("Schedule Date", "")))
            next_t_match = re.search(r"(\d{1,2}):(\d{2})", str(next_row.get("Title Start Time", "")))

            if next_d_match and next_t_match:
                nd, nm, ny = map(int, next_d_match.groups())
                nhh, nmm = map(int, next_t_match.groups())
                next_dt = datetime(ny, nm, nd, nhh, nmm, 0, tzinfo=COT)
                if next_dt > start_dt:
                    stop_dt = next_dt

        if not stop_dt:
            stop_dt = start_dt + timedelta(minutes=30)

        prog = ET.Element("programme", {
            "start": format_xmltv_date(start_dt),
            "stop": format_xmltv_date(stop_dt),
            "channel": CHANNEL_ID
        })

        title = ET.SubElement(prog, "title", {"lang": "es"})
        title.text = str(row.get("Title Name", "")).strip()

        if pd.notna(row.get("Episode Name")) and str(row.get("Episode Name")).strip():
            sub_title = ET.SubElement(prog, "sub-title", {"lang": "es"})
            sub_title.text = str(row["Episode Name"]).strip()

        if pd.notna(row.get("Title Synopsis")) and str(row.get("Title Synopsis")).strip():
            desc = ET.SubElement(prog, "desc", {"lang": "es"})
            desc.text = str(row["Title Synopsis"]).strip()

        new_programmes.append(prog)

    # Fusionar sin duplicados
    existing_starts = {p.attrib.get("start") for p in root.findall("programme") if p.attrib.get("channel") == CHANNEL_ID}
    for np in new_programmes:
        if np.attrib.get("start") not in existing_starts:
            root.append(np)

    # Purgar eventos de más de 15 días y canales viejos
    cutoff_date = datetime.now(COT) - timedelta(days=RETENTION_DAYS)
    for p in list(root.findall("programme")):
        if p.attrib.get("channel") != CHANNEL_ID:
            root.remove(p)
            continue
        stop_dt = parse_xmltv_date(p.attrib.get("stop", ""))
        if stop_dt and stop_dt < cutoff_date:
            root.remove(p)

    # Ordenar cronológicamente
    sorted_progs = sorted(
        root.findall("programme"),
        key=lambda x: parse_xmltv_date(x.attrib.get("start", "")) or datetime.min.replace(tzinfo=COT)
    )

    for p in list(root.findall("programme")):
        root.remove(p)
    for p in sorted_progs:
        root.append(p)

    ET.indent(root, space="  ", level=0)
    ET.ElementTree(root).write(XML_OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    print(f"[OK] {XML_OUTPUT_FILE} actualizado y guardado correctamente.")

if __name__ == "__main__":
    xls = download_panregional_xls()
    update_epg_xml(xls)
