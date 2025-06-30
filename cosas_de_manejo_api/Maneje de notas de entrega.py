import csv
import json
import requests
import os
import sys
import webbrowser
from datetime import date
from collections import defaultdict

#que te chupe un pingo porque es en localhost
API_KEY = "853b8fac3341fda"
API_SECRET = "2a38fe394557601"

# Nombre de los archivos a generar
archivo_crudo = "crudo.csv" #para el crudo de las OT
archivo_Series = "series_frappe.csv" #para el listado de series de frappe
archivo_productos = "productos.csv" #para el listado de productos
archivo_ONT = "columnas_extraidasONT.csv" #columnas de extraccion de datos ONT
archivo_DECO = "columnas_extraidasDeco.csv" #columnas de extraccion de datos Decos
archivo_consumibles = "reporte_consumibles.csv" #cables usados

FRAPPE_API = "http://localhost/api/resource"


def borrar_archivos_csv():
    for archivo in [archivo_crudo, archivo_Series, archivo_productos, archivo_ONT, archivo_DECO, archivo_consumibles]:
        if os.path.exists(archivo):
            os.remove(archivo)
            print(f"🧹 Archivo eliminado: {archivo}")

def hay_que_revisar(only_boolean=False):
    def tiene_pendientes(archivo):
        with open(archivo, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for fila in reader:
                if fila.get("Transferencia?", "").strip() or fila.get("No recibido", "").strip():
                    return True
        return False

    pendientes_ont = tiene_pendientes(archivo_ONT)
    pendientes_deco = tiene_pendientes(archivo_DECO)

    if only_boolean:
        return pendientes_ont or pendientes_deco

    hay_pendientes = False

    if pendientes_ont:
        print(f"⚠️ Aún hay equipos en '{archivo_ONT}' que requieren transferencia o no fueron recibidos.")
        hay_pendientes = True

    if pendientes_deco:
        print(f"⚠️ Aún hay equipos en '{archivo_DECO}' que requieren transferencia o no fueron recibidos.")
        hay_pendientes = True

    return hay_pendientes


def obtenerCrudo():
    url_csv = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSjcvuqJyiIwCYdSkAT7WGHYnGFvti3BaiswovWWWMgSTwdbVXKZQU1KrLXeiT5wJSpoqSEP9IuJ9V6/pub?gid=549752266&single=true&output=csv"

    print("Descargando CSV...")
    resp = requests.get(url_csv)
    resp.raise_for_status()

    lineas = resp.content.decode("utf-8").splitlines()

    print("Filtrando filas que no hayan sido consumidas...")
    reader = csv.reader(lineas, delimiter=',')
    filas_filtradas = [
        fila for fila in reader
        if len(fila) > 25 and fila[4].strip() == "" and (
            fila[23].strip() == "EXITOSA" or
            fila[23].strip() == "CONTINGENCIA / SUSPENDIDA"
        )
    ]

    if not filas_filtradas:
        print("🚫 No hay filas pendientes de consumo. Terminando ejecución.")
        sys.exit(0)

    with open(archivo_crudo, "w", newline='', encoding="utf-8") as f_out:
        writer = csv.writer(f_out, delimiter=',')
        writer.writerows(filas_filtradas)

    print(f"✅ Filtrado completo. {len(filas_filtradas)} filas guardadas en '{archivo_crudo}'")


def obtenerSeriesFrappe():
    print("Consultando series en Frappe...")
    
    FRAPPE_API_serie = FRAPPE_API+"/Serial%20No"
    HEADERS = {
        "Authorization": f"token {API_KEY}:{API_SECRET}"
    }

    campos = [
        "name", "docstatus", "item_code", "item_name",
        "warehouse", "creation", "modified", "_user_tags"
    ]

    fields_param = json.dumps(campos)  # ✅ Convierte a JSON válido
    url = f"{FRAPPE_API_serie}?fields={fields_param}&limit_page_length=1000"

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



def obtenerProductos():
    print("Descargando productos desde Frappe...")

    FRAPPE_API_prod = FRAPPE_API+"/Item"
    HEADERS = {
        "Authorization": f"token {API_KEY}:{API_SECRET}"
    }

    campos = ["item_code", "item_name"]
    fields_param = json.dumps(campos)
    url = f"{FRAPPE_API_prod}?fields={fields_param}&limit_page_length=1000"

    try:
        response = requests.get(url, headers=HEADERS)
        response.raise_for_status()
        data = response.json().get("data", [])

        with open(archivo_productos, "w", newline='', encoding="utf-8") as f_out:
            writer = csv.writer(f_out)
            writer.writerow(["item_code", "item_name"])
            for prod in data:
                writer.writerow([
                    prod.get("item_code", "").strip(),
                    prod.get("item_name", "").strip()
                ])

        print(f"✅ {len(data)} productos guardados en 'productos.csv'")
    except requests.exceptions.RequestException as e:
        print(f"❌ Error al descargar productos: {e}")


#--------------------------------------------------------------

def extraer_columnas_ont():
    archivo_entrada = archivo_crudo 
    archivo_salida = archivo_ONT

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
    with open(archivo_Series, newline='', encoding="utf-8") as f_series:
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
    archivo_entrada = archivo_crudo
    archivo_salida = archivo_DECO

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
    with open(archivo_Series, newline='', encoding="utf-8") as f_series:
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


def extraer_consumibles():
    archivo_salida = archivo_consumibles
    item_code_fijo = "20710003"
    item_code_extra1 = "110100372"
    item_code_extra2 = "320500065"
    
    # 1. Cargar diccionario de productos
    productos = {}
    with open(archivo_productos, newline='', encoding="utf-8") as f_prod:
        reader = csv.DictReader(f_prod)
        for row in reader:
            codigo = row["item_code"].strip()
            nombre = row["item_name"].strip()
            productos[codigo] = nombre

    # 2. Contar ocurrencias por (item_code, warehouse)
    conteo = defaultdict(lambda: defaultdict(int))  # conteo[item_code][warehouse] = cantidad

    # 🔸 Consumibles desde archivo_crudo (columna N y O)
    with open(archivo_crudo, newline='', encoding="utf-8") as f_crudo:
        reader = csv.reader(f_crudo)
        for fila in reader:
            if len(fila) > 14:
                campo_cable = fila[13].strip()  # columna N
                campo_consumible = fila[14].strip()  # columna O
                cargo_ot_base = fila[3].strip()  # columna D
                warehouse = f"{cargo_ot_base} - QPS" if cargo_ot_base else ""

                # 🔹 Cables (de columna N)
                if " - " in campo_cable:
                    item_code = campo_cable.split(" - ")[0].strip()
                    if item_code:
                        conteo[item_code][warehouse] += 1

                # 🔹 Cantidades para item_code fijo (de columna O)
                if campo_consumible:
                    try:
                        cantidad = int(campo_consumible)
                        conteo[item_code_fijo][warehouse] += cantidad
                    except ValueError:
                        print(f"⚠️ Valor no numérico ignorado en columna O: {campo_consumible}")

    # 🔸 Serializados desde archivo_DECO
    seriales_por_almacen = defaultdict(int)
    with open(archivo_DECO, newline='', encoding="utf-8") as f_deco:
        reader = csv.DictReader(f_deco)
        for fila in reader:
            serie = fila.get("Consumir", "").strip()
            almacen = fila.get("Cargó OT", "").strip()
            if serie and almacen:
                warehouse = f"{almacen} - QPS"
                seriales_por_almacen[warehouse] += 1

    for warehouse, cantidad in seriales_por_almacen.items():
        conteo[item_code_extra1][warehouse] += cantidad
        conteo[item_code_extra2][warehouse] += cantidad * 2

    # 3. Generar archivo de salida
    filas_totales = 0
    with open(archivo_salida, "w", newline='', encoding="utf-8") as f_out:
        writer = csv.writer(f_out)
        writer.writerow(["item_code", "item_name", "warehouse", "qty"])

        for item_code, almacenes in conteo.items():
            item_name = productos.get(item_code, "❓ DESCONOCIDO")
            for warehouse, qty in almacenes.items():
                writer.writerow([item_code, item_name, warehouse, qty])
                filas_totales += 1

    print(f"✅ Archivo '{archivo_salida}' generado con {filas_totales} filas.")



# ------------------------------------------------


def transferirPendientes():
    archivos = ["columnas_extraidasONT.csv", "columnas_extraidasDeco.csv"]
    url_frappe = FRAPPE_API + "/Stock%20Entry"
    headers = {
        "Authorization": f"token {API_KEY}:{API_SECRET}",
        "Content-Type": "application/json"
    }
    total_transferencias = 0

    for archivo in archivos:
        with open(archivo, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            encabezados = reader.fieldnames
            filas_actualizadas = []

            for fila in reader:
                transferencia = fila.get("Transferencia?", "").strip()
                serie = fila.get("ONT (modem fibra)", "").strip() or fila.get("Deco", "").strip()
                transferido = False

                if transferencia and serie:
                    partes = [p.strip() for p in transferencia.split(",")]
                    if len(partes) == 2:
                        origen, destino = partes
                        payload = {
                            "stock_entry_type": "Material Transfer",
                            "purpose": "Material Transfer",
                            "docstatus": 1,
                            "items": [{
                                "item_code": "",
                                "qty": 1,
                                "serial_no": serie,
                                "s_warehouse": origen,
                                "t_warehouse": destino
                            }]
                        }

                        # Buscar item_code desde series_frappe
                        with open("series_frappe.csv", newline='', encoding="utf-8") as f_series:
                            reader_series = csv.DictReader(f_series)
                            for row in reader_series:
                                if row.get("name", "").strip() == serie:
                                    payload["items"][0]["item_code"] = row.get("item_code", "").strip()
                                    break

                        print(f"🔁 Transfiriendo {serie} de {origen} a {destino}...")
                        response = requests.post(url_frappe, json=payload, headers=headers)
                        if response.status_code == 200:
                            print(f"✅ Transferencia realizada para {serie}")
                            total_transferencias += 1
                            transferido = True
                        else:
                            print(f"❌ Error en transferencia de {serie}: {response.status_code} {response.text}")

                # Si fue transferido correctamente, limpiar la columna
                if transferido:
                    fila["Transferencia?"] = ""
                filas_actualizadas.append(fila)

        # Sobrescribir el archivo con los cambios
        with open(archivo, "w", newline='', encoding="utf-8") as f_out:
            writer = csv.DictWriter(f_out, fieldnames=encabezados)
            writer.writeheader()
            writer.writerows(filas_actualizadas)

    if total_transferencias == 0:
        print("✅ No hay transferencias pendientes.")
    else:
        print(f"🏁 Total de transferencias realizadas: {total_transferencias}")


# ------------------------------------------------
def deliveryNote():
    if hay_que_revisar():
        print("🚫 No se generó la nota de entrega. Revisá los archivos antes de continuar.")
        return

    FRAPPE_API_delivery = FRAPPE_API+"/Delivery%20Note"
    COMMENT_API = FRAPPE_API+"/Communication"
    HEADERS = {
        "Authorization": f"token {API_KEY}:{API_SECRET}",
        "Content-Type": "application/json"
    }

    cliente = "Instalado"
    archivos_consumir = [archivo_ONT, archivo_DECO]

    def texto_a_html(texto):
        return texto.replace('\\n', '<br>').replace('\\t', '&emsp;')

    def agregar_comentario(nombre_doc, comentario_html, remitente="aguida@qpsrl.com.ar"):
        payload = {
            "doctype": "Communication",
            "communication_type": "Comment",
            "comment_type": "Comment",
            "reference_doctype": "Delivery Note",
            "reference_name": nombre_doc,
            "content": f"<div>{comentario_html}</div>",
            "sender": remitente
        }

        response = requests.post(COMMENT_API, headers=HEADERS, json=payload)
        if response.status_code == 200:
            print("🗨️ Comentario agregado correctamente.")
        else:
            print("❌ Error al agregar comentario:")
            print(response.status_code, response.text)

    def comentarios_en_tabla_html():
        comentarios = defaultdict(list)
        with open(archivo_crudo, newline='', encoding="utf-8") as f:
            reader = csv.reader(f)
            for fila in reader:
                if len(fila) >= 15:
                    almacen = fila[3].strip()  # columna D
                    if not almacen:
                        continue
                    fila_html = "".join(f"<td>{celda.strip()}</td>" for celda in fila[:15])
                    comentarios[f"{almacen} - QPS"].append(f"<tr>{fila_html}</tr>")
        return comentarios

    comentarios_por_almacen = comentarios_en_tabla_html()

    dic_series = {}
    with open(archivo_Series, newline='', encoding="utf-8") as f_series:
        reader = csv.reader(f_series)
        next(reader)
        for fila in reader:
            if len(fila) > 5:
                serie = fila[1].strip()
                item_code = fila[3].strip()
                item_name = fila[4].strip()
                dic_series[serie] = {
                    "item_code": item_code,
                    "item_name": item_name
                }

    agrupados = defaultdict(lambda: defaultdict(lambda: {
        "qty": 0,
        "seriales": []
    }))
    seriales_usados = set()
    for archivo in archivos_consumir:
        with open(archivo, newline='', encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for fila in reader:
                serie = fila.get("Consumir", "").strip()
                almacen = fila.get("Cargó OT", "").strip()
                if not serie or not almacen:
                    continue
                if serie in seriales_usados:
                    print(f"⚠️ Serie duplicada ignorada: {serie}")
                    continue
                if serie not in dic_series:
                    print(f"⚠️ Serie no encontrada en {archivo_Series}: {serie}")
                    continue
                info = dic_series[serie]
                key = (info["item_code"], info["item_name"])
                agrupados[almacen][key]["qty"] += 1
                agrupados[almacen][key]["seriales"].append(serie)
                seriales_usados.add(serie)

    consumibles_por_almacen = defaultdict(list)
    with open(archivo_consumibles, newline='', encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            item_code = row["item_code"].strip()
            item_name = row["item_name"].strip()
            warehouse = row["warehouse"].strip()
            try:
                qty = int(row["qty"])
                if qty > 0:
                    consumibles_por_almacen[warehouse].append({
                        "item_code": item_code,
                        "item_name": item_name,
                        "qty": qty,
                        "warehouse": warehouse
                    })
            except ValueError:
                print(f"⚠️ Cantidad inválida en consumibles: {row['qty']}")

    todos_los_almacenes = set(agrupados.keys()) | set(consumibles_por_almacen.keys())
    for almacen in todos_los_almacenes:
        items = []

        for (item_code, item_name), datos in agrupados.get(almacen, {}).items():
            items.append({
                "item_code": item_code,
                "item_name": item_name,
                "qty": datos["qty"],
                "warehouse": almacen,
                "serial_no": "\n".join(datos["seriales"])
            })

        for prod in consumibles_por_almacen.get(almacen, []):
            items.append({
                "item_code": prod["item_code"],
                "item_name": prod["item_name"],
                "qty": prod["qty"],
                "warehouse": almacen
            })

        if not items:
            continue

        filas_html = comentarios_por_almacen.get(almacen, [])
        comentario_html = ""
        if filas_html:
            comentario_html = (
                "<b>Detalle de equipos cargados:</b><br>"
                "<table border='1' cellpadding='4' cellspacing='0'>"
                f"{''.join(filas_html)}"
                "</table>"
            )

        payload = {
            "customer": cliente,
            "posting_date": str(date.today()),
            "set_warehouse": almacen,
            "items": items,
            "remarks": comentario_html
        }

        print(f"📦 Enviando nota de entrega para '{almacen}' con {len(items)} ítems...")
        response = requests.post(FRAPPE_API_delivery, headers=HEADERS, json=payload)

        if response.status_code == 200:
            docname = response.json()["data"]["name"]
            print(f"✅ Nota de entrega creada: {docname}")
            if comentario_html:
                agregar_comentario(docname, comentario_html)
        else:
            print(f"❌ Error al crear nota de entrega para {almacen}:")
            # print(response.status_code, response.text)


def abrirFotosNoRecibidos():
    archivos = ["columnas_extraidasONT.csv", "columnas_extraidasDeco.csv"]
    campo_fotos = "foto (no recibido)"
    urls = set()

    for archivo in archivos:
        try:
            with open(archivo, newline='', encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for fila in reader:
                    fotos_raw = fila.get(campo_fotos, "").strip()
                    if fotos_raw:
                        for link in fotos_raw.split("\n"):
                            url = link.strip()
                            if url.startswith("http"):
                                urls.add(url)
        except Exception as e:
            print(f"⚠️ No se pudo procesar {archivo}: {e}")

    if not urls:
        return

    print(f"Abriendo {len(urls)} fotos en el navegador...")
    for url in urls:
        webbrowser.open_new_tab(url)

def ultimaOTconsumida():
    archivo = archivo_crudo
    ultima_fila = None
    with open(archivo, newline='', encoding="utf-8") as f:
        reader = csv.reader(f)
        for fila in reader:
            ultima_fila = fila
    if ultima_fila and len(ultima_fila) > 5:
        print()
        print()
        print(f"Última OT del {archivo}: {ultima_fila[5]}")
        print()
        print()
    else:
        print(f"No se pudo leer la OT en la última fila de '{archivo}'.")



def main():
    os.system("clear")
    
    # Limpieza previa
    borrar_archivos_csv()

    #obtencion de datos
    obtenerCrudo()
    obtenerSeriesFrappe()
    obtenerProductos()

    #procesamiento
    extraer_columnas_ont()
    extraer_columnas_deco()
    extraer_consumibles()

    #transferencias
    transferirPendientes()
    
    #si tiene seguro que abre, sino no
    abrirFotosNoRecibidos()

    # Deliverys
    print()
    print()
    deliveryNote()
    
    if not hay_que_revisar(only_boolean=True):
        ultimaOTconsumida()
        borrar_archivos_csv()  # Limpieza final solo si todo salió bien

main()
