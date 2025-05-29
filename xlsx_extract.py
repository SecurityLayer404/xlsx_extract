import zipfile
import xml.etree.ElementTree as ET
import sys
import pandas as pd
import openpyxl
import re

def extract_core_metadata(zip_path):
    with zipfile.ZipFile(zip_path, 'r') as z:
        if 'docProps/core.xml' not in z.namelist():
            return None

        with z.open('docProps/core.xml') as core:
            tree = ET.parse(core)
            root = tree.getroot()

            ns = {
                'dc': 'http://purl.org/dc/elements/1.1/',
                'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                'dcterms': 'http://purl.org/dc/terms/'
            }

            metadata = {}
            for tag, label in [
                ('dc:creator', 'Autor'),
                ('cp:lastModifiedBy', 'Última modificación por'),
                ('dcterms:created', 'Fecha de creación'),
                ('dcterms:modified', 'Fecha de modificación')
            ]:
                el = root.find(tag, ns)
                metadata[label] = el.text if el is not None else 'No encontrado'
            return metadata

def search_keywords_and_users(workbook_path, keywords):
    matches = []
    usernames = set()
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    value = cell.value.strip()
                    # Buscar palabras clave
                    for kw in keywords:
                        if re.search(rf"\b{kw}\b", value, re.IGNORECASE):
                            matches.append((sheet.title, cell.coordinate, value))
                    # Generar usernames si parece un nombre completo
                    if re.match(r'^[a-zA-Z]+ [a-zA-Z]+$', value):
                        parts = value.lower().split()
                        if len(parts) == 2:
                            uname = parts[0][0] + parts[1]
                            usernames.add(uname)
    return matches, sorted(usernames)

def main(xlsx_path):
    print("📄 Extrayendo metadatos desde:", xlsx_path)
    print("-" * 60)

    try:
        metadata = extract_core_metadata(xlsx_path)
        if metadata:
            for k, v in metadata.items():
                print(f"{k:30}: {v}")
        else:
            print("No se encontraron metadatos en core.xml")
    except Exception as e:
        print(f"[!] Error al extraer metadatos XML: {e}")

    print("\n🔍 Buscando palabras clave y posibles usernames:")
    keywords = ['password', 'user', 'username', 'flag', 'admin', 'secret']
    try:
        matches, usernames = search_keywords_and_users(xlsx_path, keywords)

        if matches:
            print("\n📌 Coincidencias encontradas:")
            for sheet, coord, value in matches:
                print(f"[{sheet}] {coord}: {value}")
        else:
            print("No se encontraron coincidencias relevantes.")

        if usernames:
            print(f"\n✅ {len(usernames)} posibles usuarios extraídos:")
            for u in usernames:
                print(f" - {u}")
            with open("users.txt", "w") as f:
                for u in usernames:
                    f.write(u + "\n")
            print("\n📁 Archivo 'users.txt' generado con posibles usernames.")
        else:
            print("No se detectaron posibles nombres de usuario.")

    except Exception as e:
        print(f"[!] Error al procesar el archivo: {e}")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Uso: python3 xlsx_metadata_user_extractor.py archivo.xlsx")
    else:
        main(sys.argv[1])
