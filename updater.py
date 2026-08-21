import os
import re
import json
from datetime import datetime, timedelta, timezone
import xml.etree.ElementTree as ET
import pandas as pd
import requests

# -------------------------------------------------------------
# CONFIGURACIÓN GENERAL Y CANALES
# -------------------------------------------------------------
XML_OUTPUT_FILE = "CNLA_EPG.xml"
RETENTION_DAYS = 15

AUTH_URL = "https://epg.tapkit.warnermedia.com/api/security/oauth/token"
BASE_DAILY_URL = "https://epg.tapkit.warnermedia.com/api/daily/shows?feedId={feed_id}&format=xls"

FEEDS_CONFIG = [
    {
        "feed_id": "CNLA_PAN",
        "channel_id": "CNLA_PAN.co",
        "channel_name": "Cartoon Network Panregional",
        "lang": "es",
        "tz": timezone(timedelta(hours=-5)),
        "tz_str": "-0500"
    },
    {
        "feed_id": "CNLA_BR",
        "channel_id": "CNLA_BR.br",
        "channel_name": "Cartoon Network Brasil",
        "lang": "pt",
        "tz": timezone(timedelta(hours=-3)),
        "tz_str": "-0300"
    }
]

COMMON_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es",
    "Origin": "https://epg.tapkit.warnermedia.com",
    "Referer": "https://epg.tapkit.warnermedia.com/login"
}

def login_and_get_token():
    email = os.environ.get("TAPKIT_EMAIL", "").strip()
    password = os.environ.get("TAPKIT_PASSWORD", "").strip()

    if not email or not password:
        raise Exception("Faltan TAPKIT_EMAIL o TAPKIT_PASSWORD en los Secrets de GitHub.")

    payload = {
        "grant_type": "password",
        "scope": "any",
        "username": email,
        "password": password
    }

    res = requests.post(AUTH_URL, data=payload, headers=COMMON_HEADERS, timeout=30)
    if res.status_code != 200:
        raise Exception(f"Error en login. Código HTTP: {res.status_code} | Respuesta: {res.text}")

    data = res.json()
    token = data.get("access_token")
    if not token:
        raise Exception(f"No se encontró 'access_token' en la respuesta: {data}")

    print("[OK] Sesión iniciada y token obtenido exitosamente.")
    return token

def download_feed_xls(token, feed_id):
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "es",
        "Authorization": f"Bearer {token}",
        "Referer": "https://epg.tapkit.warnermedia.com/epg/networks/2"
    })

    url = BASE_DAILY_URL.format(feed_id=feed_id)
    res = session.get(url, timeout=60)
    res.raise_for_status()

    xls_path = f"{feed_id}_latest.xls"
    with open(xls_path, "wb") as f:
        f.write(res.content)

    print(f"[OK] XLS descargado exitosamente para {feed_id}.")
    return xls_path

def sanitize_and_parse_xml(file_path):
    if not os.path.exists(file_path):
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network MultiFeed"})

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
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network MultiFeed"})

def parse_xmltv_date(date_str, tz_info):
    if not date_str:
        return None
    clean_str = re.sub(r"[^\d\s\+\-]", "", str(date_str).strip())
    match = re.match(r"^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", clean_str)
    if not match:
        return None

    year, month, day, hour, minute, second = map(int, match.groups())
    hour, minute, second = min(hour, 23), min(minute, 59), min(second, 59)
    return datetime(year, month, day, hour, minute, second, tzinfo=tz_info)

def format_xmltv_date(dt, tz_str):
    return dt.strftime(f"%Y%m%d%H%M%S {tz_str}")

def load_excel_schedule(file_path):
    try:
        df = pd.read_excel(file_path, header=1, sheet_name=0)
    except Exception:
        df = pd.read_html(file_path, header=1)[0]
    return df

def process_feed(root, feed_cfg, xls_path):
    channel_id = feed_cfg["channel_id"]
    channel_name = feed_cfg["channel_name"]
    lang = feed_cfg["lang"]
    tz = feed_cfg["tz"]
    tz_str = feed_cfg["tz_str"]

    # Registrar el canal si aún no existe
    existing_channels = [ch for ch in root.findall("channel") if ch.attrib.get("id") == channel_id]
    if not existing_channels:
        ch_node = ET.SubElement(root, "channel", {"id": channel_id})
        disp = ET.SubElement(ch_node, "display-name")
        disp.text = channel_name

    df = load_excel_schedule(xls_path)
    cols = {str(c).strip(): c for c in df.columns}
    col_date = cols.get("Schedule Date", df.columns[0])
    col_time = cols.get("Title Start Time", df.columns[1])
    col_title = cols.get("Title Name", df.columns[2])
    col_ep = cols.get("Episode Name", df.columns[3] if len(df.columns) > 3 else None)
    col_desc = cols.get("Title Synopsis", df.columns[4] if len(df.columns) > 4 else None)
    col_ep_desc = cols.get("Episode Synopsis", None)

    first_date_raw = str(df.iloc[0].get(col_date, "")).strip()
    match_init = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", first_date_raw)
    if not match_init:
        raise Exception(f"No se pudo detectar la fecha inicial de {channel_id}.")

    d0, m0, y0 = map(int, match_init.groups())
    cycle_start = datetime(y0, m0, d0, 6, 0, 0, tzinfo=tz)
    cycle_end = cycle_start + timedelta(days=1)

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

        event_dt = datetime(y, m, d, hh, mm, 0, tzinfo=tz)
        if (d == d0 and m == m0 and y == y0) and hh < 6:
            event_dt += timedelta(days=1)

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

    raw_events = sorted(raw_events, key=lambda x: x["start"])
    new_programmes = []
    total_events = len(raw_events)

    for i in range(total_events):
        ev = raw_events[i]
        start_dt = ev["start"]
        stop_dt = raw_events[i + 1]["start"] if i + 1 < total_events else cycle_end

        prog = ET.Element("programme", {
            "start": format_xmltv_date(start_dt, tz_str),
            "stop": format_xmltv_date(stop_dt, tz_str),
            "channel": channel_id
        })

        title = ET.SubElement(prog, "title", {"lang": lang})
        title.text = ev["title"]

        if ev["sub_title"]:
            sub_title = ET.SubElement(prog, "sub-title", {"lang": lang})
            sub_title.text = ev["sub_title"]

        if ev["desc"]:
            desc = ET.SubElement(prog, "desc", {"lang": lang})
            desc.text = ev["desc"]

        new_programmes.append(prog)

    existing_starts = {p.attrib.get("start") for p in root.findall("programme") if p.attrib.get("channel") == channel_id}
    for np in new_programmes:
        if np.attrib.get("start") not in existing_starts:
            root.append(np)

    print(f"[{channel_id}] Eventos procesados hoy: {len(new_programmes)}")

def main():
    token = login_and_get_token()
    root = sanitize_and_parse_xml(XML_OUTPUT_FILE)

    active_channel_ids = {cfg["channel_id"] for cfg in FEEDS_CONFIG}
    for ch in list(root.findall("channel")):
        if ch.attrib.get("id") not in active_channel_ids:
            root.remove(ch)

    # Descarga y procesa cada feed
    for feed_cfg in FEEDS_CONFIG:
        xls_path = download_feed_xls(token, feed_cfg["feed_id"])
        process_feed(root, feed_cfg, xls_path)

    # Limpieza por retención histórica (15 días)
    cutoff_date = datetime.now(timezone.utc) - timedelta(days=RETENTION_DAYS)
    for p in list(root.findall("programme")):
        ch_id = p.attrib.get("channel")
        if ch_id not in active_channel_ids:
            root.remove(p)
            continue
        cfg = next((c for c in FEEDS_CONFIG if c["channel_id"] == ch_id), FEEDS_CONFIG[0])
        stop_dt = parse_xmltv_date(p.attrib.get("stop", ""), cfg["tz"])
        if stop_dt and stop_dt < cutoff_date:
            root.remove(p)

    # Reordenar elementos XMLTV
    sorted_progs = sorted(
        root.findall("programme"),
        key=lambda x: x.attrib.get("start", "")
    )
    for p in list(root.findall("programme")):
        root.remove(p)
    for p in sorted_progs:
        root.append(p)

    ET.indent(root, space="  ", level=0)
    ET.ElementTree(root).write(XML_OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    print(f"[OK] XML multi-feed actualizado con éxito. Total acumulado: {len(sorted_progs)} eventos.")

if __name__ == "__main__":
    main()
