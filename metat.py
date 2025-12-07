import xml.etree.ElementTree as ET
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import re
import json
from lxml import etree
import requests
from uri_template import variable
from bs4 import BeautifulSoup
import urllib


# === Namespaces ===
namespaces = {
    'gmd': "http://www.isotc211.org/2005/gmd",
    'gco': "http://www.isotc211.org/2005/gco",
    'srv': "http://www.isotc211.org/2005/srv"
}

# === Lizenz-Mapping ===
def normalize_license_text(text):
    return re.sub(r'[^a-z0-9]', '', text.lower())

LICENSE_MAP = {
    normalize_license_text("datenlizenz deutschland – zero – version 2.0"): "https://www.govdata.de/dl-de/zero-2-0",
    normalize_license_text("datenlizenz deutschland – namensnennung – version 2.0"): "https://www.govdata.de/dl-de/by-2-0",
    normalize_license_text("cc-by 4.0"): "https://creativecommons.org/licenses/by/4.0/",
    normalize_license_text("es gelten keine bedingungen"): "https://www.govdata.de/dl-de/zero-2-0"
}


def map_license_url(freetext):
    if not freetext:
        return None
    freetext = freetext.strip()
    try:
        parsed = json.loads(freetext)
        if isinstance(parsed, dict) and 'url' in parsed:
            return parsed['url']
    except json.JSONDecodeError:
        pass
    normalized = normalize_license_text(freetext)
    return LICENSE_MAP.get(normalized, "manuell prüfen")

# === URL-Prüfung ===
def check_url_reachable(url):
    if not url:
        return False
    headers = {
        'User-Agent': 'Mozilla/5.0'
    }
    try:
        response = requests.head(url, headers=headers, allow_redirects=True, timeout=5)
        if response.status_code < 400:
            return True
        response = requests.get(url, headers=headers, allow_redirects=True, timeout=5)
        return response.status_code < 400
    except Exception:
        return False

# === DCAT-AP-konforme URL-Unterscheidung ===
def get_dcat_urls_strict(root):
    download_url = None
    access_url = None
    transfer_options = root.findall('.//gmd:transferOptions//gmd:onLine//gmd:CI_OnlineResource', namespaces)
    for res in transfer_options:
        url_el = res.find('.//gmd:URL', namespaces)
        if url_el is not None and url_el.text:
            url = url_el.text.strip()
            url_lower = url.lower()
            is_direct_file = any(url_lower.endswith(ext) for ext in ['.zip', '.csv', '.gml', '.xml', '.geojson', '.json', '?'])
            if is_direct_file:
                if not download_url:
                    download_url = url

            else:
                if not access_url:
                    access_url = url
    return download_url, access_url

def get_text(root, xpath):
    el = root.find(xpath, namespaces)
    return el.text.strip() if el is not None and el.text else None



# === INSPIRE-/ISO-Prüfung ===
def is_inspire_conform(root):
    std = get_text(root, './/gmd:metadataStandardName/gco:CharacterString')
    if not std:
        return False
    std = std.lower()
    return any(key in std for key in ['iso 19115', 'iso19115', 'iso 19119', 'iso19119', 'inspire'])

# === XPath-Helfer ===
def get_text_debug(root, xpath):
    el = root.find(xpath, namespaces)
    if el is not None and el.text:
        print(f"[DEBUG] Found text for {xpath}: {el.text.strip()}")
        return el.text.strip()
    else:
        print(f"[DEBUG] Nothing found for {xpath}")
        return None

# === FAIR-Indikatoren ===
def check_rda_i1_02m_etree(file_path):
    try:
        tree = etree.parse(file_path)
        root = tree.getroot()
        ns_set = {elem.tag.split("}")[0][1:] for elem in root.iter() if elem.tag.startswith("{")}
        return "ja" if ns_set & {
            "http://www.w3.org/ns/dcat#", "http://schema.org/",
            "http://www.w3.org/2004/02/skos/core#", "http://purl.org/dc/terms/",
            "http://www.w3.org/1999/02/22-rdf-syntax-ns#", "http://www.w3.org/2002/07/owl#"
        } else "nein"
    except:
        return "Fehler"

def check_rda_i2_01m_etree(file_path):
    try:
        tree = etree.parse(file_path)
        root = tree.getroot()
        ns_set = {elem.tag.split("}")[0][1:] for elem in root.iter() if elem.tag.startswith("{")}
        return "ja" if ns_set & {
            "http://www.isotc211.org/2005/gmd", "http://www.opengis.net/gml",
            "http://www.isotc211.org/2005/gco", "http://www.w3.org/ns/dcat#",
            "http://purl.org/dc/terms/"
        } else "nein"
    except:
        return "Fehler"

def check_rda_r1_3_01d(format_text):
    if not format_text:
        return "nein"
    valid = ['application/x-esri-shapefile','application/geo+json','application/gml+xml','text/csv','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','text/xml','RDF','OGC:WFS','OGC:WMS','application/json']
    return "ja" if any(v in format_text for v in valid) else "nein"

def check_rda_a1_1_01d(d, z):
    return "ja" if any((u and u.startswith(('http','https','ftp'))) for u in [d,z]) else "nein"

def check_rda_a1_04d(d, z):
    return check_rda_a1_1_01d(d, z)

# Pop-up Fenster für manuelle Eingaben (3 pages: Kategorien -> Bundesland -> Optionen)
def popup(title: str, geo_desc: str):
    class Option(ttk.Frame):
        def __init__(self, parent, text, variable):
            super().__init__(parent)
            ttk.Label(self, text=text).pack()
            ttk.Radiobutton(self, text='ja', variable=variable, value=True).pack()
            ttk.Radiobutton(self, text='nein', variable=variable, value=False).pack()
            variable.set(True)

    class Checkbox(ttk.Checkbutton):
        def __init__(self, parent, text, variable):
            super().__init__(parent, text=text, variable=variable, offvalue=False, onvalue=True)
            self.pack(anchor="w")

    # ----- Root-Fenster -----
    root = tk.Tk()
    root.title('Felder manuell eingeben')
    root.geometry('360x480')

    ttk.Label(root, text=title).pack(pady=6)


    container = ttk.Frame(root)
    container.pack(side='top', fill='both', expand=True)

    page0 = ttk.Frame(container)
    page1 = ttk.Frame(container)
    page2 = ttk.Frame(container)

    for p in (page0, page1, page2):
        p.place(relx=0, rely=0, relwidth=1, relheight=1)

    # ===== PAGE 0: Kategorien =====
    ttk.Label(page0, text='Kategorien auswählen').pack(pady=4)

    kategorien = {
        'Gebiet': tk.BooleanVar(),
        'Gebäude': tk.BooleanVar(),
        'Klima': tk.BooleanVar(),
        'Landwirtschaft': tk.BooleanVar(),
        'Bildung': tk.BooleanVar(),
        'Gesundheit': tk.BooleanVar(),
        'Wirtschaft': tk.BooleanVar(),
        'Bevölkerung': tk.BooleanVar(),
        'Sicherheit': tk.BooleanVar(),
        'Umwelt': tk.BooleanVar(),
        'Energie': tk.BooleanVar(),
        'Technologie': tk.BooleanVar(),
        'Transport': tk.BooleanVar(),
        'anderes': tk.BooleanVar()
    }
    for k in kategorien:
        Checkbox(page0, k, kategorien[k])

    # ===== PAGE 1: Wähle Bundesland (Combobox) =====
    ttk.Label(page1, text='Wähle Bundesland').pack(pady=8)

    BUNDESLAENDER = [
        "Schleswig-Holstein","Hamburg","Niedersachsen","Bremen","Nordrhein-Westfalen",
        "Hessen","Rheinland-Pfalz","Baden-Württemberg","Bayern","Saarland","Berlin",
        "Brandenburg","Mecklenburg-Vorpommern","Sachsen","Sachsen-Anhalt","Thüringen"
    ]
    bundesland = tk.StringVar(value=geo_desc)
    cb = ttk.Combobox(page1, values=BUNDESLAENDER, textvariable=bundesland, state="readonly", width=28)
    cb.pack(pady=6)

    # ===== PAGE 2 =====
    synthetische_daten = tk.BooleanVar()
    ohne_zahlung = tk.BooleanVar()
    ohne_registrierung = tk.BooleanVar()

    Option(page2, "enthält synthetische Daten", synthetische_daten).pack(pady=4)
    Option(page2, "ist zugänglich ohne Zahlung", ohne_zahlung).pack(pady=4)
    Option(page2, "ist zugänglich ohne Registrierung", ohne_registrierung).pack(pady=4)

    rahmen = ttk.Frame(page2)
    rahmen.pack(pady=6)
    ttk.Label(rahmen, text='Erstellenart').pack()
    erstellenart = tk.StringVar(value='amtlich')
    ttk.Radiobutton(rahmen, text='amtlich', variable=erstellenart, value='amtlich').pack()
    ttk.Radiobutton(rahmen, text='privat', variable=erstellenart, value='privat').pack()
    ttk.Radiobutton(rahmen, text='crowdsourced', variable=erstellenart, value='crowdsourced').pack()

    # ----- Navigation -----
    pages = [page0, page1, page2]
    idx = {'i': 0}
    pages[0].lift()

    def sammeln_und_schliessen():
        # Kategorien zusammensetzen
        category = '; '.join([k for k, v in kategorien.items() if v.get()])
        data = {
            'Kategorie': category,
            'Bundesland': bundesland.get(),
            'enthält synthetische Daten': 'ja' if synthetische_daten.get() else 'nein',
            'ist zugänglich ohne Zahlung': 'ja' if ohne_zahlung.get() else 'nein',
            'ist zugänglich ohne Registrierung': 'ja' if ohne_registrierung.get() else 'nein',
            'Erstellenart': erstellenart.get()
        }
        root.quit()
        root.destroy()
        return data

    result_holder = {'data': None}

    def next_click():
        if idx['i'] < len(pages) - 1:
            idx['i'] += 1
            pages[idx['i']].lift()
            if idx['i'] == len(pages) - 1:
                next_button.config(text='Fertig')
        else:
            result_holder['data'] = sammeln_und_schliessen()

    def prev_click():
        if idx['i'] > 0:
            idx['i'] -= 1
            pages[idx['i']].lift()
            next_button.config(text='Weiter')

    btn_frame = ttk.Frame(root)
    btn_frame.pack(side="bottom", anchor="e", pady=6, padx=6)
    back_button = ttk.Button(btn_frame, text="Zurück", command=prev_click)
    back_button.pack(side="left", padx=(0,6))
    next_button = ttk.Button(btn_frame, text="Weiter", command=next_click)
    next_button.pack(side="left")

    root.mainloop()
    return result_holder['data']

# === Scrape opengeodata.nrw.de for download files ===

def get_url_extensions_nrw(url):
    headers = {
        'Accept': 'application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    }
    try:
        web = requests.get(url, headers=headers, allow_redirects=True)
    except:
        return {'Zugriffs-URL nicht erreichbar'}
    if web.status_code != 200:
        return {'Zugriffs-URL nicht erreichbar'}

    soup = BeautifulSoup(web.text, "xml")
    files = soup.find_all('files')[1]

    # download_urls = []
    # for file in files:
    #     if file.name == None:
    #         continue
    #     download_urls.append(urllib.parse.urljoin(web.url,file['name']))

    return [file['name'] for file in files if file.name]

# === Return download files as urls

def get_download_urls_nrw(url, files):
    for f in files:
        if f == 'Zugriffs-URL nicht erreichbar':
            return []
    download_urls = []
    for file in files:
        download_urls.append(urllib.parse.urljoin(url,file))

    return download_urls
# === Format / Service to Recommended DCAT Entry (IANA “Media Types” Vokabular) ===

MEDIA_TYPES = {
    r"\bShapefile \b": ["application/x-esri-shapefile"],
    r"\bGeoPackage\b|\bGPKG\b": ["application/geopackage+sqlite3", "application/geopackage"],
    r"\bGML\b": ["application/gml+xml"],
    r"\bGeoJSON \b": ["application/geo+json"],
    r"\bKML\b": ["application/vnd.google-earth.kml+xml"],
    r"\bCSV\b": ["text/csv"],
    r"\bNetCDF\b": ["application/x-netcdf"],
    r"\bTIFF\b|\bGeoTIFF\b": ["image/tiff", "image/geotiff"],
    r"\bJPEG2000\b|\bjp2\b": ["image/jp2"],
    r"\bPDF\b": ["application/pdf"],
    r"\bZIP\b": ["application/zip"],
    r"\bXML\b": ["text/xml", "application/xml"],
    r"\bWMS\b": ["OGC:WMS", "application/xml"],
    r"\bWFS\b": ["OGC:WFS", "application/xml"],
    r"\batom\b|inspire download service": ["application/atom+xml"],
    r"\b(gdb|file geodatabase|geodatabase)\b": ["application/x-esri-filegdb"],
    r"\bsqlite\b(?!.*geopackage)": ["application/vnd.sqlite3"],
    r"\bjson\b(?!.*geojson)": ["application/json"],
    r"\b(xlsx|excel)\b": ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
}

def recommended_dcat_entry(format_service: str) -> str:
    if not format_service:
        return ""
    text = format_service.lower()
    for pattern, media_list in MEDIA_TYPES.items():
        if re.search(pattern, text):
            return " | ".join(media_list)   # nhiều lựa chọn -> nối bằng " | "
    return format_service  # fallback nếu không khớp

# === Einzelner Metadatensatz ===
def extract_metadata(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    if not is_inspire_conform(root):
        return None
    geo_raw = get_text(root, './/gmd:extent//gmd:EX_Extent//gmd:description/gco:CharacterString')
    if not geo_raw:
        geo_raw = get_text(root, './/gmd:EX_Extent/gmd:description/gco:CharacterString')
    if not geo_raw:
        geo_raw = get_text(root, './/gmd:extent//gmd:description/gco:CharacterString')
    if not geo_raw:
        geo_raw = get_text(root, './/gmd:EX_GeographicDescription//gmd:MD_Identifier//gmd:code/gco:CharacterString')
    if not geo_raw:
        geo_raw = get_text(root, './/gmd:country/gco:CharacterString')

# === Bundesland-Mapping basierend auf 2-stelligem AGS-Code ===
    BUNDESLAND_SCHLUESSEL = {
        "01": "Schleswig-Holstein",
        "02": "Hamburg",
        "03": "Niedersachsen",
        "04": "Bremen",
        "05": "Nordrhein-Westfalen",
        "06": "Hessen",
        "07": "Rheinland-Pfalz",
        "08": "Baden-Württemberg",
        "09": "Bayern",
        "10": "Saarland",
        "11": "Berlin",
        "12": "Brandenburg",
        "13": "Mecklenburg-Vorpommern",
        "14": "Sachsen",
        "15": "Sachsen-Anhalt",
        "16": "Thüringen"
    }

    geo_desc = None


    if geo_raw:
        if geo_raw.strip() == "276" or "deutschland" in geo_raw.lower():
            geo_desc = "Deutschland"
        elif re.fullmatch(r'\d{12}', geo_raw.strip()):
            bl_code = geo_raw.strip()[:2]
            bl_name = BUNDESLAND_SCHLUESSEL.get(bl_code)
            if bl_name:
                geo_desc = bl_name
            else:
                geo_desc = geo_raw  # Fallback
        else:
            geo_desc = geo_raw
    else:
        geo_desc = None



    # === Lizenz korrekt aus allen möglichen Constraints extrahieren ===
    license_texts = root.findall('.//gmd:resourceConstraints//gmd:otherConstraints/gco:CharacterString', namespaces)
    license_url = None
    for lt in license_texts:
        if lt.text:
            license_url = map_license_url(lt.text.strip())
            if license_url and license_url != "manuell prüfen":
                break
    if not license_url:
        license_url = "manuell prüfen"

    download_url, access_url = get_dcat_urls_strict(root)
    download_files = []
    download_urls = []

    # === Prüfe Download-URL erreichbar
    if not download_url:
        download_files = get_url_extensions_nrw(access_url)
        download_urls = get_download_urls_nrw(access_url, download_files)
        download_url = '; '.join(download_urls)
    elif not check_url_reachable(download_url):
        download_url += " (Bitte manuell angeben, URL nicht erreichbar)"

    # Zugriffs-URL lassen wir unangetastet, auch wenn sie evtl. nicht erreichbar ist


    file_id = get_text(root, './/gmd:fileIdentifier/gco:CharacterString')
    identifier = get_text(root, './/srv:identifier/gco:CharacterString')
    title = get_text(root, './/gmd:title/gco:CharacterString')

    manual_data = popup(title, geo_desc)

    data = {
        'Übernommen von Appsmith': '',
        'Metadatensatz_ID': identifier,
        'Datensatz_ID': file_id,
        'Titel': title,
        'Beschreibung': get_text(root, './/gmd:abstract/gco:CharacterString'),
        'Kategorie': manual_data.get('Kategorie'),
        'enthält synthetische Daten': manual_data.get('enthält synthetische Daten'),
        'ist zugänglich ohne Zahlung': manual_data.get('ist zugänglich ohne Zahlung'),
        'ist zugänglich ohne Registrierung': manual_data.get('ist zugänglich ohne Registrierung'),
        'Erstellenart': manual_data.get('Erstellenart'),
        'Geographische Beschreibung': manual_data.get('Bundesland'),
        'Lizenz': license_url,
        'Herausgeber': get_text(root, './/gmd:pointOfContact//gmd:organisationName/gco:CharacterString'),
        'Kontakt E-Mail': get_text(root, './/gmd:electronicMailAddress/gco:CharacterString'),
        'Download-URL': download_url,
        'Zugriffs-URL': access_url,
        'Metadatenstandard': get_text(root, './/gmd:metadataStandardName/gco:CharacterString'),
        'Metadatenstandardversion': get_text(root, './/gmd:metadataStandardVersion/gco:CharacterString'),
        'Veröffentlichungsdatum': get_text(root, './/gmd:date//gco:DateTime'),
        'Letzte Aktualisierung': get_text(root, './/gmd:dateStamp/gco:Date'),
        'Erstellungsdatum des Metadatensatzes': get_text(root, './/gmd:dateStamp/gco:Date'),
        'Format': recommended_dcat_entry(get_text(root, './/gmd:distributionFormat//gmd:name/gco:CharacterString'))
    }

    # === FAIR-Erweiterung ===
    data.update({
        'RDA-F1-01M': 'ja' if file_id else 'nein', #changed
        'RDA-F1-01D': 'ja' if identifier else 'nein', #changed
        'RDA-F1-02M': 'ja' if file_id and file_id.startswith('http') else 'nein',
        'RDA-F1-02D': 'ja' if identifier and identifier.startswith('http') else 'nein',
        'RDA-F2-01M': 'ja' if all([data['Titel'], data['Beschreibung'], data['Format'], license_url]) else 'nein',
        'RDA-F3-01M': 'ja' if file_id or access_url else 'nein', #changed
        'RDA-A1-01M': 'ja' if download_url or access_url else 'nein',
        'RDA-A1-02M': 'ja' if data['Kontakt E-Mail'] or download_url or access_url else 'nein', #changed
        'RDA-A1-02D': 'ja' if data['Kontakt E-Mail'] or download_url or access_url else 'nein',
        'RDA-A1-04M': 'ja' if download_url.startswith('http') else 'nein',
        'RDA-A1-04D': check_rda_a1_04d(download_url, access_url),
        'RDA-A1.1-01M': 'ja' if download_url.startswith('http') else 'nein',
        'RDA-A1.1-01D': check_rda_a1_1_01d(download_url, access_url),
        'RDA-I1-01M': 'ja' if data['Metadatenstandard'] else 'nein', #changed
        'RDA-I1-02M': check_rda_i1_02m_etree(file_path),
        'RDA-I2-01M': check_rda_i2_01m_etree(file_path),
        'RDA-R1.1-01M': 'ja' if license_url else 'nein',
        'RDA-R1.3-01M': 'ja' if data['Metadatenstandard'] else 'nein',
        'RDA-R1.3-01D': check_rda_r1_3_01d(data['Format']),
        'RDA-R1.3-02M': 'ja' if any(x in str(data.get('Metadatenstandard','')).lower() #changed
                 for x in ['iso', 'iso/ts', 'rdf', 'owl', 'xsd', 'dcat']) else 'nein',
        'Eintragsdatum': datetime.now().strftime('%Y-%m-%d'),
        'Keywords': '', 'Kommentar': '', 'Person': '' #changed
    })

    entries = []
    if download_urls:
        for file, url in zip(download_files, download_urls):
            data['Titel'], data['Download-URL'] = file, url
            entries.append(data.copy())
    else:
        entries.append(data)
    return entries

# === Benutzerinput & Excel-Ausgabe ===
def get_user_input():
    root = tk.Tk()
    root.withdraw()
    xml_dir = filedialog.askdirectory(title="XML-Verzeichnis auswählen")
    if not xml_dir:
        return None, None
    excel_file = filedialog.asksaveasfilename(
        title="Excel-Datei speichern unter", defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx")]
    )
    root.destroy()
    return xml_dir, excel_file

# === Hauptfunktion ===
def main():
    xml_dir, excel_file = get_user_input()
    if not xml_dir or not excel_file:
        return
    files = [os.path.join(xml_dir, f) for f in os.listdir(xml_dir) if f.endswith('.xml')]
    entries = []
    for f in files:
        if f:
            data = extract_metadata(f)
            for d in data:
                entries.append(d)

    if not entries:
        print("Keine gültigen INSPIRE-/ISO19115/19119-Metadaten gefunden.")
        return
    df = pd.DataFrame(entries)
    df.to_excel(excel_file, index=False)
    print(f"{len(df)} Datensätze gespeichert in: {excel_file}")

if __name__ == "__main__":
    main()
