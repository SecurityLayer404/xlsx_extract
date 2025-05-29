
# xlsx_extract

`xlsx_extract.py` es una herramienta de análisis forense y de reconocimiento pensada para CTFs y auditorías de seguridad. Permite extraer metadatos, buscar palabras clave relevantes y generar automáticamente una lista de posibles usuarios a partir de archivos `.xlsx` (Excel).

## 🚀 Características

- Extrae metadatos del documento (.xlsx) como autor, editor y fechas de creación/modificación.
- Busca automáticamente términos clave típicos de entornos corporativos: `password`, `admin`, `flag`, etc.
- Detecta posibles nombres completos y genera usernames (formato inicial + apellido).
- Guarda los usernames en un archivo `users.txt` listo para usar en herramientas como `kerbrute`, `GetNPUsers.py`, etc.

## 🛠 Requisitos

- Python 3.6+
- Bibliotecas:

```bash
pip install openpyxl pandas
```

## 📦 Instalación y uso

1. Cloná el repositorio o descargá el script:

```bash
git clone https://github.com/SecurityLayer404/xlsx_extract.git
cd xlsx_extract
```

2. Ejecutá el script sobre tu archivo `.xlsx`:

```bash
python3 xlsx_extract.py archivo.xlsx
```

3. El script mostrará resultados en consola y generará (si aplica) un archivo `users.txt`.

## 📁 Ejemplo de salida

```bash
📄 Extrayendo metadatos desde: employee_status.xlsx
Autor                          : L1lith
Última modificación por        : xyan1d3
Fecha de creación              : 2025-05-27T22:01:00Z
Fecha de modificación          : 2025-05-28T07:42:21Z

🔍 Buscando palabras clave y posibles usernames:
[Sheet1] B7: admin password
✅ 5 posibles usuarios extraídos:
 - jsmith
 - dport
 - mlopez

📁 Archivo 'users.txt' generado con posibles usernames.
```

## ⚠️ Uso ético

Esta herramienta está diseñada para actividades educativas, legales y autorizadas. Su uso en entornos no permitidos puede ser ilegal.

---

