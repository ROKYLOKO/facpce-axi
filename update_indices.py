"""
update_indices.py
Busca el Excel más reciente de índices FACPCE (Res. JG 539/18),
lo descarga, lo parsea y genera indices.json.

Uso local:
    pip install requests beautifulsoup4 openpyxl
    python update_indices.py
"""

import requests
from bs4 import BeautifulSoup
import re
import openpyxl
import json
import io
import sys
from datetime import datetime

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}
XLSX_PATTERN = re.compile(
    r'https?://(?:www\.)?facpce\.org\.ar/wp-content/uploads/\d{4}/\d{2}/Indice-FACPCE[^"\'>\s]+\.xlsx',
    re.IGNORECASE,
)
SEARCH_URLS = [
    "https://www.facpce.org.ar/",
    "https://www.facpce.org.ar/indices/",
    "https://www.facpce.org.ar/area-tecnica/",
]
# Fallback: URL conocida (actualizar si la búsqueda falla)
FALLBACK_URL = (
    "https://www.facpce.org.ar/wp-content/uploads/2026/06/"
    "Indice-FACPCE-Res.-JG-539-18-2026-05-1.xlsx"
)

MONTHS_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}
MONTH_ABBR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
               "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def find_xlsx_url() -> str | None:
    """Escanea páginas de FACPCE buscando el link al Excel más reciente."""
    found = []
    for page in SEARCH_URLS:
        try:
            r = requests.get(page, headers=HEADERS, timeout=15)
            r.raise_for_status()
            found.extend(XLSX_PATTERN.findall(r.text))
        except Exception as e:
            print(f"  [warn] {page}: {e}")
    if found:
        # Ordenar para quedarnos con el más reciente (mayor fecha en URL)
        found = sorted(set(found))
        return found[-1]
    return None


def parse_date_cell(cell) -> datetime | None:
    """Intenta convertir una celda a fecha."""
    if isinstance(cell, datetime):
        return cell
    if not isinstance(cell, str):
        return None
    s = cell.strip().lower()
    # mm/yyyy o mm-yyyy
    m = re.match(r"^(\d{1,2})[/\-](\d{4})$", s)
    if m:
        return datetime(int(m.group(2)), int(m.group(1)), 1)
    # yyyy-mm
    m = re.match(r"^(\d{4})[/\-](\d{1,2})$", s)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), 1)
    # "enero 2016" / "enero-2016"
    m = re.match(r"^([a-záéíóú]+)[\s\-](\d{4})$", s)
    if m and m.group(1) in MONTHS_ES:
        return datetime(int(m.group(2)), MONTHS_ES[m.group(1)], 1)
    # "ene-16" / "ene 16"
    abbr_map = {a.lower(): i + 1 for i, a in enumerate(MONTH_ABBR)}
    m = re.match(r"^([a-z]{3})[\s\-](\d{2,4})$", s)
    if m and m.group(1) in abbr_map:
        yr = int(m.group(2))
        if yr < 100:
            yr += 2000
        return datetime(yr, abbr_map[m.group(1)], 1)
    return None


def parse_xlsx(content: bytes) -> list[dict]:
    """Parsea el contenido binario del Excel y devuelve lista de {period, label, index}."""
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    ws = wb.active

    raw_rows = []
    for row in ws.iter_rows(values_only=True):
        if not any(c is not None for c in row):
            continue
        raw_rows.append(row)
    wb.close()

    data = []
    for row in raw_rows:
        date_val = None
        index_val = None
        for cell in row:
            if cell is None:
                continue
            d = parse_date_cell(cell)
            if d is not None:
                date_val = d
            if isinstance(cell, (int, float)) and 50 < cell < 9_999_999:
                # Evitar seriales de fechas Excel (circa 40000-50000)
                if not (40_000 < cell < 50_000):
                    index_val = float(cell)
        if date_val and index_val and 2010 <= date_val.year <= 2035:
            data.append({
                "period": date_val.strftime("%Y-%m"),
                "label": f"{MONTH_ABBR[date_val.month - 1]} {date_val.year}",
                "index": round(index_val, 6),
            })

    # Ordenar y deduplicar
    data.sort(key=lambda x: x["period"])
    seen: set[str] = set()
    dedup = []
    for r in data:
        if r["period"] not in seen:
            seen.add(r["period"])
            dedup.append(r)
    return dedup


def main():
    print("=" * 50)
    print("Actualizador de Índice FACPCE AXI")
    print("=" * 50)

    print("\n[1/3] Buscando archivo más reciente en facpce.org.ar...")
    url = find_xlsx_url()
    if url:
        print(f"      Encontrado: {url}")
    else:
        print(f"      No encontrado. Usando URL de respaldo:\n      {FALLBACK_URL}")
        url = FALLBACK_URL

    print("\n[2/3] Descargando Excel...")
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        print(f"      Descargado ({len(r.content) / 1024:.1f} KB)")
    except Exception as e:
        print(f"      ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    print("\n[3/3] Parseando y generando indices.json...")
    data = parse_xlsx(r.content)
    if not data:
        print("      ERROR: No se pudieron extraer datos.", file=sys.stderr)
        sys.exit(1)

    output = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_url": url,
        "count": len(data),
        "data": data,
    }
    with open("indices.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✓ {len(data)} períodos guardados en indices.json")
    print(f"  Rango: {data[0]['label']} → {data[-1]['label']}")
    print()


if __name__ == "__main__":
    main()
