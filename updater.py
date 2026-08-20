import os
import re
from datetime import datetime, timedelta, timezone
import xml.etree.ElementTree as ET
import pandas as pd
from playwright.sync_api import sync_playwright

# -------------------------------------------------------------
# CONFIGURACIÓN
# -------------------------------------------------------------
XML_OUTPUT_FILE = "CNLA_EPG.xml"
CHANNEL_ID = "CNLA_PAN.co"
CHANNEL_NAME = "Cartoon Network Panregional"
RETENTION_DAYS = 15
COT = timezone(timedelta(hours=-5))

def download_panregional_xls():
    email = os.environ.get("TAPKIT_EMAIL")
    password = os.environ.get("TAPKIT_PASSWORD")
    download_dir = os.path.abspath("downloads")
    os.makedirs(download_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # 1. Login
        page.goto("https://epg.tapkit.warnermedia.com/login", wait_until="domcontentloaded")
        page.wait_for_selector('input[placeholder*="correo"], input[formcontrolname="email"], input[type="email"]', timeout=45000)

        email_input = page.locator('input[placeholder*="correo"], input[formcontrolname="email"], input[type="email"]').first
        pass_input = page.locator('input[placeholder*="contrase"], input[formcontrolname="password"], input[type="password"]').first

        email_input.fill(email)
        pass_input.fill(password)

        page.locator('button:has-text("Ingresar"), button[type="submit"]').first.click()
        page.wait_for_url("**/epg/**", timeout=45000)

        # 2. Ir a Cartoon Network
        page.goto("https://epg.tapkit.warnermedia.com/epg/networks/2", wait_until="domcontentloaded")
        page.wait_for_selector("table.dailygridstable, app-network-resources", timeout=45000)

        # 3. Descargar fila Panregional HD
        row = page.locator('tr:has-text("Cartoon Network Panregional")')
        with page.expect_download(timeout=60000) as download_info:
            row.locator("button").first.click()

        download = download_info.value
        xls_path = os.path.join(download_dir, "CNLA_PAN_latest.xls")
        download.save_as(xls_path)
        browser.close()
        return xls_path

def parse_xmltv_date(date_str):
    match = re.match(r"^(\d{14})\s*([+-]\d{4})?", date_str.strip())
    if not match:
        return None
    dt_part, tz_part = match.groups()
    dt = datetime.strptime(dt_part, "%Y%m%d%H%M%S")
    if tz_part:
        tz_h, tz_m = int(tz_part[:3]), int(tz_part[0] + tz_part[3:])
        return dt.replace(tzinfo=timezone(timedelta(hours=tz_h, minutes=tz_m)))
    return dt.replace(tzinfo=COT)

def format_xmltv_date(dt):
    return dt.strftime("%Y%m%d%H%M%S -0500")

def update_epg_xml(xls_path):
    df = pd.read_excel(xls_path)

    # 1. Cargar el archivo CNLA_EPG.xml existente
    if os.path.exists(XML_OUTPUT_FILE):
        try:
            tree = ET.parse(XML_OUTPUT_FILE)
            root = tree.getroot()
        except Exception:
            root = ET.Element("tv", {"generator-info-name": "EPG Auto Updater"})
    else:
        root = ET.Element("tv", {"generator-info-name": "EPG Auto Updater"})

    # Asegurar el nodo del canal Panregional y limpiar el nodo de México si existe
    for ch in list(root.findall("channel")):
        if ch.attrib.get("id") != CHANNEL_ID:
            root.remove(ch)

    if not any(ch.attrib.get("id") == CHANNEL_ID for ch in root.findall("channel")):
        ch_node = ET.SubElement(root, "channel", {"id": CHANNEL_ID})
        disp = ET.SubElement(ch_node, "display-name")
        disp.text = CHANNEL_NAME

    # 2. Procesar programas desde el Excel descargado
    new_programmes = []
    for i in range(len(df)):
        row = df.iloc[i]
        date_str = str(row["Schedule Date"]).strip()
        time_str = str(row["Title Start Time"]).strip()

        try:
            start_dt = datetime.strptime(f"{date_str} {time_str}", "%d-%m-%Y %H:%M").replace(tzinfo=COT)
        except Exception:
            continue

        if i + 1 < len(df):
            next_row = df.iloc[i + 1]
            try:
                stop_dt = datetime.strptime(f"{str(next_row['Schedule Date']).strip()} {str(next_row['Title Start Time']).strip()}", "%d-%m-%Y %H:%M").replace(tzinfo=COT)
            except Exception:
                stop_dt = start_dt + timedelta(minutes=30)
        else:
            stop_dt = start_dt + timedelta(minutes=30)

        prog = ET.Element("programme", {
            "start": format_xmltv_date(start_dt),
            "stop": format_xmltv_date(stop_dt),
            "channel": CHANNEL_ID
        })

        title = ET.SubElement(prog, "title", {"lang": "es"})
        title.text = str(row.get("Title Name", ""))

        if pd.notna(row.get("Episode Name")):
            sub_title = ET.SubElement(prog, "sub-title", {"lang": "es"})
            sub_title.text = str(row["Episode Name"])

        if pd.notna(row.get("Title Synopsis")):
            desc = ET.SubElement(prog, "desc", {"lang": "es"})
            desc.text = str(row["Title Synopsis"])

        new_programmes.append(prog)

    # 3. Unir evitando duplicados
    existing_starts = {p.attrib.get("start") for p in root.findall("programme") if p.attrib.get("channel") == CHANNEL_ID}
    for np in new_programmes:
        if np.attrib.get("start") not in existing_starts:
            root.append(np)

    # 4. Purgar eventos antiguos (> 15 días) y descartar programas no panregionales
    cutoff_date = datetime.now(COT) - timedelta(days=RETENTION_DAYS)
    for p in list(root.findall("programme")):
        if p.attrib.get("channel") != CHANNEL_ID:
            root.remove(p)
            continue
        stop_dt = parse_xmltv_date(p.attrib.get("stop", ""))
        if stop_dt and stop_dt < cutoff_date:
            root.remove(p)

    # Ordenar eventos cronológicamente
    programmes = sorted(root.findall("programme"), key=lambda x: x.attrib.get("start", ""))
    for p in list(root.findall("programme")):
        root.remove(p)
    for p in programmes:
        root.append(p)

    # 5. Guardar en CNLA_EPG.xml
    ET.indent(root, space="  ", level=0)
    ET.ElementTree(root).write(XML_OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    print(f"[OK] {XML_OUTPUT_FILE} actualizado exitosamente.")

if __name__ == "__main__":
    xls = download_panregional_xls()
    update_epg_xml(xls)
