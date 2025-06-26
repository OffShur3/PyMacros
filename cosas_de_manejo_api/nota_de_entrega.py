import csv
import requests
import sys
import os
from io import StringIO
from datetime import date

FRAPPE_API = "http://localhost/api/resource/Delivery%20Note"
API_KEY = "853b8fac3341fda"
API_SECRET = "2a38fe394557601"

url_csv = "https://docs.google.com/spreadsheets/d/17Qkf859KKPN9YGH_q2cuMLMNhmnbUTu9v3bsLVBUul4/pub?gid=1213551365&single=true&output=csv"
cliente = sys.argv[1]
almacen = sys.argv[2]
comentario = sys.argv[3]


HEADERS = {
    "Authorization": f"token {API_KEY}:{API_SECRET}",
    "Content-Type": "application/json"
}

EXPECTED_FIELDS = ["item_code", "qty", "warehouse", "serial_no"]

os.system('clear')

def texto_a_html(texto):
    return texto.replace('\\n', '<br>').replace('\\t', '&emsp;')

comentario = texto_a_html(comentario) 

def buscar_fila_de_encabezado_en_lineas(lineas):
    for i, line in enumerate(lineas):
        lowered = line.strip().lower()
        if "item code" in lowered or "item_code" in lowered:
            return i
    return None

def leer_items_csv_desde_url(url):
    response = requests.get(url)
    response.raise_for_status()
    contenido = response.text.splitlines()

    header_line = buscar_fila_de_encabezado_en_lineas(contenido)
    if header_line is None:
        print("❌ No se encontró la fila de encabezados válida.")
        return []

    csv_text = "\n".join(contenido[header_line:])
    f = StringIO(csv_text)

    reader = csv.DictReader(f)
    print("🧩 Encabezados:", reader.fieldnames)
    items = []
    for i, row in enumerate(reader, 1):
        item_code = row.get("item_code") or row.get("Item Code")
        qty_raw = row.get("qty") or row.get("Quantity")
        warehouse = row.get("warehouse") or row.get("From Warehouse")
        serial_no = row.get("serial_no") or row.get("Serial No")

        if not item_code or not qty_raw:
            print(f"⚠️  Fila {i} ignorada: falta item_code o qty")
            continue

        if isinstance(qty_raw, str) and qty_raw.strip().lower() in ["qty", "quantity", ""]:
            print(f"⚠️  Fila {i} ignorada: qty no es numérico ({qty_raw})")
            continue

        try:
            qty = float(qty_raw)
        except ValueError:
            print(f"⚠️  Fila {i} ignorada: qty no convertible a número ({qty_raw})")
            continue

        item = {
            "item_code": item_code.strip(),
            "qty": qty,
            "warehouse": warehouse.strip() if warehouse else "Almacén - QPS"
        }

        if serial_no and serial_no.strip():
            seriales = [s.strip() for s in serial_no.split("\n") if s.strip()]
            item["serial_no"] = "\n".join(seriales)

        items.append(item)

    return items

def agregar_comentario(nombre_doc, comentario_html, remitente="aguida@qpsrl.com.ar"):
    """Agrega un comentario a una nota de entrega ya creada"""
    payload = {
        "doctype": "Communication",
        "communication_type": "Comment",
        "comment_type": "Comment",
        "reference_doctype": "Delivery Note",
        "reference_name": nombre_doc,
        "content": f"<div>{comentario_html}</div>",
        "sender": remitente
    }

    response = requests.post(
        "http://localhost/api/resource/Communication",
        headers=HEADERS,
        json=payload
    )

    if response.status_code == 200:
        print("🗨️ Comentario agregado correctamente.")
    else:
        print("❌ Error al agregar comentario:")
        print(response.status_code, response.text)

def crear_delivery_note(cliente, items, warehouse, comentario=""):
    payload = {
        "customer": cliente,
        "posting_date": str(date.today()),
        "set_warehouse": warehouse,
        "items": items,
        "remarks": comentario
    }

    print("📦 Enviando datos a Frappe...")
    response = requests.post(FRAPPE_API, headers=HEADERS, json=payload)

    if response.status_code == 200:
        docname = response.json()["data"]["name"]
        print(f"✅ Nota de entrega creada exitosamente: {docname}")
        agregar_comentario(docname, comentario)
        os.system("xdg-open http://localhost/desk#Form/Delivery%20Note/"+docname)
    else:
        print("❌ Error al crear la nota de entrega:")
        print(response.status_code, response.text)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Uso: python3 nota_de_entrega.py <cliente> <almacen> <comentario>")
        sys.exit(1)

    items = leer_items_csv_desde_url(url_csv)
    if not items:
        print("❌ No se encontraron ítems válidos en el CSV.")
        sys.exit(1)

    crear_delivery_note(cliente, items, almacen, comentario)
