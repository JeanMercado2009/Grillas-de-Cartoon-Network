import os
import re
import glob
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
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        context = browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        # 1. Login
        page.goto("https://epg.tapkit.warnermedia.com/login", wait_until="domcontentloaded")
        page.wait_for_selector('input[type="email"], input[formcontrolname="email"], input[placeholder*="correo"]', timeout=45000)

        email_input = page.locator('input[type="email"], input[formcontrolname="email"], input[placeholder*="correo"]').first
        pass_input = page.locator('input[type="password"], input[formcontrolname="password"], input[placeholder*="contrase"]').first

        email_input.fill(email)
        pass_input.fill(password)

        page.locator('button:has-text("Ingresar"), button[type="submit"]').first.click()
        page.wait_for_url("**/epg/**", timeout=45000)

        # 2. Navegar a Cartoon Network
        page.goto("https://epg.tapkit.warnermedia.com/epg/networks/2", wait_until="domcontentloaded")
        page.wait_for_selector("table.dailygridstable, .custom-sidebar", timeout=45000)

        # 3. Localizar el botón exacto de Panregional
        row = page.locator('tr:has-text("Cartoon Network Panregional")')
        download_btn = row.locator('button.feed-button, td.cdk-column-download button, button:has(i.fa-file-alt)').first

        with page.expect_download(timeout=60000) as download_info:
            download_btn.click()

        download = download_info.value
        xls_path = os.path.join(download_dir, "CNLA_PAN_latest.xls")
        download.save_as(xls_path)
        browser.close()
        return xls_path

def sanitize_and_parse_xml(file_path):
    """
    Limpia texto plano suelto (como 'JUEVES', 'VIERNES') y etiquetas huérfanas
    para evitar que ElementTree lance ParseError.
    """
    if not os.path.exists(file_path):
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network Panregional"})

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        # Eliminar texto suelto de días que no esté dentro de etiquetas XML
        cleaned_lines = []
        for line in content.splitlines():
            stripped = line.strip()
            # Ignorar líneas de texto suelto de días
            if stripped in ["JUEVES", "VIERNES", "SÁBADO", "SABADO", "DOMINGO", "LUNES", "MARTES", "MIÉRCOLES", "MIERCOLES"] or "PANREGIONAL" in stripped and not stripped.startswith("<"):
                continue
            cleaned_lines.append(line)

        cleaned_content = "\n".join(cleaned_lines)
        return ET.fromstring(cleaned_content)
    except Exception:
        return ET.Element("tv", {"generator-info-name": "Guia de Programacion Cartoon Network Panregional"})

def parse_xmltv_date(date_str):
    if not date_str:
        return None
    
    # Corrige errores tipográficos como '202608069194000' (15 dígitos) recortando el exceso
    clean_str = re.sub(r"[^\d\s\+\-]", "", date_str.strip())
    match = re.match(r"^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", clean_str)
    
    if not match:
        return None

    year, month, day, hour, minute, second = map(int, match.groups())
    
    # Validar rangos de hora y minuto
    hour = min(hour, 23)
    minute = min(minute, 59)
    second = min(second, 59)

    dt = datetime(year, month, day, hour, minute, second)
    return dt.replace(tzinfo=COT)

def format_xmltv_date(dt):
    return dt.strftime("%Y%m%d%H%M%S -0500")

def update_epg_xml(xls_path):
    # Cargar Excel descargado
    df = pd.read_excel(xls_path)

    # 1. Cargar y sanitizar XML existente
    root = sanitize_and_parse_xml(XML_OUTPUT_FILE)

    # Asegurar únicamente el canal CNLA_PAN.co
    for ch in list(root.findall("channel")):
        if ch.attrib.get("id") != CHANNEL_ID:
            root.remove(ch)

    if not any(ch.attrib.get("id") == CHANNEL_ID for ch in root.findall("channel")):
        ch_node = ET.SubElement(root, "channel", {"id": CHANNEL_ID})
        disp = ET.SubElement(ch_node, "display-name")
        disp.text = CHANNEL_NAME

    # 2. Parsear los programas del archivo XLS
    new_programmes = []
    total_rows = len(df)

    for i in range(total_rows):
        row = df.iloc[i]
        date_raw = str(row.get("Schedule Date", "")).strip()
        time_raw = str(row.get("Title Start Time", "")).strip()

        # Extraer fecha DD-MM-YYYY
        date_match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", date_raw)
        time_match = re.search(r"(\d{1,2}):(\d{2})", time_raw)

        if not date_match or not time_match:
            continue

        d, m, y = map(int, date_match.groups())
        hh, mm = map(int, time_match.groups())
        start_dt = datetime(y, m, d, hh, mm, 0).replace(tzinfo=COT)

        # Calcular hora de fin
        stop_dt = None
        if i + 1 < total_rows:
            next_row = df.iloc[i + 1]
            next_date_raw = str(next_row.get("Schedule Date", "")).strip()
            next_time_raw = str(next_row.get("Title Start Time", "")).strip()
            next_d_match = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", next_date_raw)
            next_t_match = re.search(r"(\d{1,2}):(\d{2})", next_time_raw)

            if next_d_match and next_t_match:
                nd, nm, ny = map(int, next_d_match.groups())
                nhh, nmm = map(int, next_t_match.groups())
                next_dt = datetime(ny, nm, nd, nhh, nmm, 0).replace(tzinfo=COT)
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

    # 3. Unir programas y evitar duplicados por hora de inicio
    existing_starts = {p.attrib.get("start") for p in root.findall("programme") if p.attrib.get("channel") == CHANNEL_ID}
    for np in new_programmes:
        if np.attrib.get("start") not in existing_starts:
            root.append(np)

    # 4. Eliminar canales que no sean Panregional y purgar > 15 días
    cutoff_date = datetime.now(COT) - timedelta(days=RETENTION_DAYS)
    for p in list(root.findall("programme")):
        if p.attrib.get("channel") != CHANNEL_ID:
            root.remove(p)
            continue
        stop_dt = parse_xmltv_date(p.attrib.get("stop", ""))
        if stop_dt and stop_dt < cutoff_date:
            root.remove(p)

    # 5. Ordenar cronológicamente
    sorted_progs = sorted(
        root.findall("programme"),
        key=lambda x: parse_xmltv_date(x.attrib.get("start", "")) or datetime.min.replace(tzinfo=COT)
    )

    for p in list(root.findall("programme")):
        root.remove(p)
    for p in sorted_progs:
        root.append(p)

    # 6. Guardar archivo XML formateado
    ET.indent(root, space="  ", level=0)
    tree = ET.ElementTree(root)
    tree.write(XML_OUTPUT_FILE, encoding="utf-8", xml_declaration=True)
    print(f"[OK] {XML_OUTPUT_FILE} procesado correctamente sin errores.")

if __name__ == "__main__":
    xls = download_panregional_xls()
    update_epg_xml(xls)
