import csv
import json
import requests


API_KEY = "853b8fac3341fda"
API_SECRET = "2a38fe394557601"

def obtenerCrudo():
    # URL del CSV desde Google Sheets
    url_csv = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSjcvuqJyiIwCYdSkAT7WGHYnGFvti3BaiswovWWWMgSTwdbVXKZQU1KrLXeiT5wJSpoqSEP9IuJ9V6/pub?gid=549752266&single=true&output=csv"

    # Nombre del archivo local
    archivo_crudo = "crudo.csv"

    # Paso 1: Descargar el archivo CSV
    print("Descargando CSV...")
    resp = requests.get(url_csv)
    resp.raise_for_status()

    # Decodificar contenido con UTF-8 explícitamente
    lineas = resp.content.decode("utf-8").splitlines()

    # Paso 2: Leer las líneas y filtrar filas con columna A vacía y estado correcto
    print("Filtrando filas que no hayan sido consumidas...")

    reader = csv.reader(lineas, delimiter=',')
    filas_filtradas = [
        fila for fila in reader
        if len(fila) > 25 and fila[4].strip() == "" and ( #fila E
            fila[23].strip() == "EXITOSA" or
            fila[23].strip() == "CONTINGENCIA / SUSPENDIDA" #fila x
        )
    ]

    # Paso 3: Sobrescribir crudo.csv con los resultados filtrados
    with open(archivo_crudo, "w", newline='', encoding="utf-8") as f_out:
        writer = csv.writer(f_out, delimiter=',')
        writer.writerows(filas_filtradas)

    print(f"✅ Filtrado completo. {len(filas_filtradas)} filas guardadas en '{archivo_crudo}'")

def obtenerSeriesFrappe():
    print("Consultando series en Frappe...")
    
    archivo_Series = "series_frappe.csv"
    FRAPPE_API = "http://localhost/api/resource/Serial%20No"
    HEADERS = {
        "Authorization": f"token {API_KEY}:{API_SECRET}"
    }

    campos = [
        "name", "docstatus", "item_code", "item_name",
        "warehouse", "creation", "modified", "_user_tags"
    ]

    fields_param = json.dumps(campos)  # ✅ Convierte a JSON válido
    url = f"{FRAPPE_API}?fields={fields_param}&limit_page_length=1000"

    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json().get("data", [])

        with open(archivo_Series, "w", newline='', encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["sr"] + campos)  # encabezados

            for i, fila in enumerate(data, 1):
                writer.writerow([
                    i,
                    fila.get("name", ""),
                    fila.get("docstatus", ""),
                    fila.get("item_code", ""),
                    fila.get("item_name", ""),
                    fila.get("warehouse", ""),
                    fila.get("creation", ""),
                    fila.get("modified", ""),
                    fila.get("_user_tags", "")
                ])

        print(f"✅ {len(data)} series guardadas en '{archivo_Series}'")

    except requests.exceptions.RequestException as e:
        print(f"❌ Error al hacer la solicitud a Frappe: {e}")


def extraer_columnas():
    archivo_entrada = "crudo.csv"
    archivo_salida = "columnas_extraidasONT.csv"
    archivo_series = "series_frappe.csv"

    encabezados = [
        "ONT (modem fibra)",
        "foto",
        "Cargó OT",
        "Origen equipo",
        "Transferencia?",
        "No recibido",
        "foto (no recibido)",
        "Consumir"
    ]

    # Leer series_frappe.csv y construir un diccionario {serie: warehouse}
    dic_ont_origen = {}
    with open(archivo_series, newline='', encoding="utf-8") as f_series:
        reader_series = csv.reader(f_series)
        next(reader_series)  # saltar encabezado
        for fila in reader_series:
            if len(fila) > 5:
                serie = fila[1].strip()
                warehouse = fila[5].strip()
                dic_ont_origen[serie] = warehouse

    # Leer crudo.csv y procesar
    with open(archivo_entrada, newline='', encoding="utf-8") as f_in:
        reader = csv.reader(f_in)
        filas_salida = []

        for fila in reader:
            if len(fila) > 9 and fila[9].strip() != "":
                ont = fila[9].strip()
                foto = fila[8].strip()  # crudo (con comas)
                cargo_ot_base = fila[3].strip()
                cargo_ot = f"{cargo_ot_base} - QPS" if cargo_ot_base else ""

                # Determinar origen equipo
                if ont in dic_ont_origen and dic_ont_origen[ont].strip():
                    origen_equipo = dic_ont_origen[ont].strip()
                elif ont not in dic_ont_origen:
                    origen_equipo = "No"
                else:
                    origen_equipo = ""

                # Transferencia (solo si origen_equipo no es "No")
                if origen_equipo and origen_equipo != cargo_ot and origen_equipo != "No":
                    transferencia = f"{origen_equipo}, {cargo_ot}"
                else:
                    transferencia = ""

                no_recibido = ""
                foto_no_recibido = ""

                if origen_equipo == "No":
                    no_recibido = ont
                    foto_no_recibido = "\n".join([f.strip() for f in foto.split(",") if f.strip()])

                # Consumir si no necesita transferencia y no tiene origen "No"
                consumir = ""
                if origen_equipo != "No" and origen_equipo == cargo_ot:
                    consumir = ont

                filas_salida.append([
                    ont,
                    foto,
                    cargo_ot,
                    origen_equipo,
                    transferencia,
                    no_recibido,
                    foto_no_recibido,
                    consumir
                ])

    with open(archivo_salida, "w", newline='', encoding="utf-8") as f_out:
        writer = csv.writer(f_out)
        writer.writerow(encabezados)
        writer.writerows(filas_salida)

    print(f"✅ Archivo '{archivo_salida}' generado con {len(filas_salida)} filas filtradas.")


def extraer_columnas_deco():
    archivo_entrada = "crudo.csv"
    archivo_salida = "columnas_extraidasDeco.csv"
    archivo_series = "series_frappe.csv"

    encabezados = [
        "Deco",
        "foto",
        "Cargó OT",
        "Origen equipo",
        "Transferencia?",
        "No recibido",
        "foto (no recibido)",
        "Consumir"
    ]

    # Leer series_frappe.csv y construir un diccionario {serie: warehouse}
    dic_deco_origen = {}
    with open(archivo_series, newline='', encoding="utf-8") as f_series:
        reader_series = csv.reader(f_series)
        next(reader_series)  # saltar encabezado
        for fila in reader_series:
            if len(fila) > 5:
                serie = fila[1].strip()
                warehouse = fila[5].strip()
                dic_deco_origen[serie] = warehouse

    # Leer crudo.csv y procesar
    with open(archivo_entrada, newline='', encoding="utf-8") as f_in:
        reader = csv.reader(f_in)
        filas_salida = []

        for fila in reader:
            if len(fila) > 10 and fila[10].strip() != "":
                deco_series = fila[10].strip().split()  # Puede haber múltiples series separados por espacio
                foto = fila[8].strip()
                cargo_ot_base = fila[3].strip()
                cargo_ot = f"{cargo_ot_base} - QPS" if cargo_ot_base else ""

                for deco in deco_series:
                    origen_equipo = "No"
                    if deco in dic_deco_origen and dic_deco_origen[deco].strip():
                        origen_equipo = dic_deco_origen[deco].strip()
                    elif deco not in dic_deco_origen:
                        origen_equipo = "No"
                    else:
                        origen_equipo = ""

                    # Transferencia (solo si origen_equipo no es "No")
                    if origen_equipo and origen_equipo != cargo_ot and origen_equipo != "No":
                        transferencia = f"{origen_equipo}, {cargo_ot}"
                    else:
                        transferencia = ""

                    no_recibido = ""
                    foto_no_recibido = ""
                    if origen_equipo == "No":
                        no_recibido = deco
                        foto_no_recibido = "\n".join([f.strip() for f in foto.split(",") if f.strip()])

                    # Consumir: si no requiere transferencia y no tiene origen "No"
                    consumir = ""
                    if origen_equipo != "No" and origen_equipo == cargo_ot:
                        consumir = deco

                    filas_salida.append([
                        deco,
                        foto,
                        cargo_ot,
                        origen_equipo,
                        transferencia,
                        no_recibido,
                        foto_no_recibido,
                        consumir
                    ])

    with open(archivo_salida, "w", newline='', encoding="utf-8") as f_out:
        writer = csv.writer(f_out)
        writer.writerow(encabezados)
        writer.writerows(filas_salida)

    print(f"✅ Archivo '{archivo_salida}' generado con {len(filas_salida)} filas.")


def main():
    obtenerCrudo()
    obtenerSeriesFrappe()
    extraer_columnas()
    extraer_columnas_deco()

main()
