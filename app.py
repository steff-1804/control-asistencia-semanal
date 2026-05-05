from flask import Flask, render_template, request, redirect, url_for
from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import quote
import os
import requests
import unicodedata

# ==========================================================
# RUTA LOCAL DEL PROYECTO EN TU COMPUTADORA
# ==========================================================

RUTA_LOCAL = r"C:\Users\MASTER\OneDrive\Documentos\Nueva carpeta\TECNICO"

if os.path.exists(RUTA_LOCAL):
    BASE_DIR = RUTA_LOCAL
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

app = Flask(__name__, template_folder=TEMPLATE_DIR)

# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

ZONA_HORARIA = ZoneInfo("America/Guayaquil")

AIRTABLE_TOKEN = os.environ.get("AIRTABLE_TOKEN")
AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID")

TABLA_PERSONAL = os.environ.get("AIRTABLE_TABLA_PERSONAL", "Personal")
TABLA_CHARLAS = os.environ.get("AIRTABLE_TABLA_CHARLAS", "Charlas")
TABLA_ASISTENCIA = os.environ.get("AIRTABLE_TABLA_ASISTENCIA", "Asistencia")


def ahora_ecuador():
    return datetime.now(ZONA_HORARIA)


def normalizar_texto(texto):
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def convertir_semana(numero_semana):
    """
    Convierte:
    1 -> Semana 1
    2 -> Semana 2
    3 -> Semana 3
    """

    numero_semana = str(numero_semana).strip()

    if not numero_semana:
        raise ValueError("Debe ingresar el número de semana.")

    if not numero_semana.isdigit():
        raise ValueError("Debe ingresar solo el número de semana. Ejemplo: 1")

    return f"Semana {numero_semana}"


def airtable_headers():
    if not AIRTABLE_TOKEN:
        raise ValueError(
            "No se encontró AIRTABLE_TOKEN. "
            "Debes configurarlo en Render o en las variables de entorno de tu computadora."
        )

    if not AIRTABLE_BASE_ID:
        raise ValueError(
            "No se encontró AIRTABLE_BASE_ID. "
            "Debes configurarlo en Render o en las variables de entorno de tu computadora."
        )

    return {
        "Authorization": f"Bearer {AIRTABLE_TOKEN}",
        "Content-Type": "application/json"
    }


def airtable_url(tabla):
    tabla_codificada = quote(tabla)
    return f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{tabla_codificada}"


def airtable_get_records(tabla, filter_formula=None):
    url = airtable_url(tabla)
    headers = airtable_headers()

    records = []
    params = {}

    if filter_formula:
        params["filterByFormula"] = filter_formula

    while True:
        response = requests.get(url, headers=headers, params=params, timeout=30)

        if response.status_code != 200:
            raise Exception(
                f"Error leyendo Airtable: {response.status_code} - {response.text}"
            )

        data = response.json()
        records.extend(data.get("records", []))

        offset = data.get("offset")
        if not offset:
            break

        params["offset"] = offset

    return records


def airtable_create_records(tabla, records):
    if not records:
        return

    url = airtable_url(tabla)
    headers = airtable_headers()

    # Airtable acepta máximo 10 registros por solicitud.
    for i in range(0, len(records), 10):
        lote = records[i:i + 10]

        payload = {
            "records": lote,
            "typecast": True
        }

        response = requests.post(url, headers=headers, json=payload, timeout=30)

        if response.status_code not in [200, 201]:
            raise Exception(
                f"Error creando registros en Airtable: {response.status_code} - {response.text}"
            )


def airtable_delete_records(tabla, record_ids):
    if not record_ids:
        return

    url = airtable_url(tabla)
    headers = airtable_headers()

    # Airtable acepta máximo 10 eliminaciones por solicitud.
    for i in range(0, len(record_ids), 10):
        lote = record_ids[i:i + 10]

        params = []
        for record_id in lote:
            params.append(("records[]", record_id))

        response = requests.delete(url, headers=headers, params=params, timeout=30)

        if response.status_code != 200:
            raise Exception(
                f"Error eliminando registros anteriores: {response.status_code} - {response.text}"
            )


def obtener_personal_por_area(area):
    area_normalizada = normalizar_texto(area)

    registros = airtable_get_records(TABLA_PERSONAL)

    personal = []

    for registro in registros:
        campos = registro.get("fields", {})

        cedula = str(campos.get("Cedula", "")).strip()
        nombre = str(campos.get("Nombre", "")).strip()
        area_persona = str(campos.get("Area", "")).strip()
        estado = str(campos.get("Estado", "ACTIVO")).strip()

        if not cedula or not nombre:
            continue

        if normalizar_texto(area_persona) != area_normalizada:
            continue

        if normalizar_texto(estado) != "ACTIVO":
            continue

        personal.append({
            "cedula": cedula,
            "nombre": nombre,
            "area": area_persona
        })

    personal.sort(key=lambda x: x["nombre"])

    return personal


def obtener_charlas_por_semana_area(semana, area):
    area_normalizada = normalizar_texto(area)

    registros = airtable_get_records(TABLA_CHARLAS)

    charlas = []

    for registro in registros:
        campos = registro.get("fields", {})

        fecha = str(campos.get("Fecha", "")).strip()
        semana_reg = str(campos.get("Semana", "")).strip()
        dia = str(campos.get("Dia", "")).strip()
        area_charla = str(campos.get("Area", "")).strip()
        charla = str(campos.get("Charla", "")).strip()

        if normalizar_texto(semana_reg) != normalizar_texto(semana):
            continue

        if normalizar_texto(area_charla) != area_normalizada:
            continue

        if not fecha or not dia or not charla:
            continue

        charlas.append({
            "fecha": fecha,
            "semana": semana_reg,
            "dia": dia,
            "area": area_charla,
            "charla": charla
        })

    orden_dias = {
        "LUNES": 1,
        "MARTES": 2,
        "MIERCOLES": 3,
        "MIÉRCOLES": 3,
        "JUEVES": 4,
        "VIERNES": 5,
        "SABADO": 6,
        "SÁBADO": 6,
        "DOMINGO": 7
    }

    charlas.sort(
        key=lambda x: (
            orden_dias.get(normalizar_texto(x["dia"]), 99),
            x["fecha"]
        )
    )

    return charlas


def eliminar_asistencia_previa(semana, area):
    """
    Evita duplicados.
    Si ya existe asistencia guardada para esa semana y área,
    la elimina y vuelve a guardar la lista completa.
    """

    registros = airtable_get_records(TABLA_ASISTENCIA)

    ids_a_eliminar = []

    for registro in registros:
        campos = registro.get("fields", {})

        semana_reg = str(campos.get("Semana", "")).strip()
        area_reg = str(campos.get("Area", "")).strip()

        if (
            normalizar_texto(semana_reg) == normalizar_texto(semana)
            and normalizar_texto(area_reg) == normalizar_texto(area)
        ):
            ids_a_eliminar.append(registro["id"])

    airtable_delete_records(TABLA_ASISTENCIA, ids_a_eliminar)


@app.route("/", methods=["GET"])
def inicio():
    return render_template("inicio.html")


@app.route("/semanal", methods=["GET", "POST"])
def semanal():
    mensaje = ""
    tipo = ""

    if request.method == "POST":
        semana_ingresada = request.form.get("semana", "").strip()
        area = request.form.get("area", "").strip()

        if not semana_ingresada or not area:
            mensaje = "Debe ingresar la semana y el área."
            tipo = "error"
            return render_template("inicio.html", mensaje=mensaje, tipo=tipo)

        try:
            semana = convertir_semana(semana_ingresada)
        except Exception as e:
            mensaje = str(e)
            tipo = "error"
            return render_template("inicio.html", mensaje=mensaje, tipo=tipo)

        try:
            personal = obtener_personal_por_area(area)
            charlas = obtener_charlas_por_semana_area(semana, area)
        except Exception as e:
            mensaje = f"Error al cargar datos desde Airtable: {e}"
            tipo = "error"
            return render_template("inicio.html", mensaje=mensaje, tipo=tipo)

        if not personal:
            mensaje = f"No se encontró personal activo para el área {area}."
            tipo = "error"
            return render_template("inicio.html", mensaje=mensaje, tipo=tipo)

        if not charlas:
            mensaje = f"No se encontraron charlas para {semana} y el área {area}."
            tipo = "error"
            return render_template("inicio.html", mensaje=mensaje, tipo=tipo)

        return render_template(
            "semanal.html",
            semana=semana,
            semana_ingresada=semana_ingresada,
            area=area,
            personal=personal,
            charlas=charlas
        )

    return redirect(url_for("inicio"))


@app.route("/guardar-semanal", methods=["POST"])
def guardar_semanal():
    semana_ingresada = request.form.get("semana", "").strip()
    area = request.form.get("area", "").strip()

    try:
        semana = convertir_semana(semana_ingresada)
    except Exception as e:
        return render_template(
            "inicio.html",
            mensaje=str(e),
            tipo="error"
        )

    try:
        personal = obtener_personal_por_area(area)
        charlas = obtener_charlas_por_semana_area(semana, area)

        if not personal or not charlas:
            return render_template(
                "inicio.html",
                mensaje="No se pudo guardar. Revise que existan personal y charlas para esa semana y área.",
                tipo="error"
            )

        eliminar_asistencia_previa(semana, area)

        hora_registro = ahora_ecuador().strftime("%H:%M:%S")

        registros_para_crear = []

        for persona in personal:
            for charla in charlas:
                campo_checkbox = f"asistencia__{persona['cedula']}__{charla['fecha']}"

                asistio = "✓" if request.form.get(campo_checkbox) == "on" else "✗"

                registros_para_crear.append({
                    "fields": {
                        "Fecha": charla["fecha"],
                        "Semana": semana,
                        "Dia": charla["dia"],
                        "Cedula": persona["cedula"],
                        "Nombre": persona["nombre"],
                        "Area": area,
                        "Charla": charla["charla"],
                        "Asistencia": asistio,
                        "HoraRegistro": hora_registro
                    }
                })

        airtable_create_records(TABLA_ASISTENCIA, registros_para_crear)

        return render_template(
            "inicio.html",
            mensaje=f"Asistencia semanal guardada correctamente para {area}, {semana}.",
            tipo="exito"
        )

    except Exception as e:
        return render_template(
            "inicio.html",
            mensaje=f"No se pudo guardar la asistencia. Error: {e}",
            tipo="error"
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)