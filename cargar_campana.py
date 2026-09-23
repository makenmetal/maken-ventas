"""Carga las 36 empresas de la campaña de correo al sistema de ventas.

Fuente: datos/campana.xlsx del pipeline de prospección (giro, municipio,
contacto) más la bitácora real de la campaña — qué se mandó, cuándo, qué rebotó
y quién contestó.

Por omisión NO escribe: imprime lo que haría. Con --subir lo manda al API.
"""
import argparse
import json
import sys
import urllib.request

import pandas as pd

API = ("https://script.google.com/macros/s/"
       "AKfycbym1-YyG7LPm2fOxSrad50vxDHHiHbKRDaJLJ5jVF-T83N_kbqN9SZJuL75pwNJ2slCYw/exec")
VENDEDOR = "Juan Pablo Lopez"
XLSX = r"C:\Users\jplop\prospeccion-maken\datos\campana.xlsx"

# Cada empresa: (clave en el Excel, fecha del 1er correo, fecha del seguimiento
# o None, desenlace). El desenlace manda sobre el Estado.
CAMPANA = [
    ("DUROSA",                  "2026-08-26", "2026-09-22", "sin_respuesta"),
    ("AURIN",                   "2026-08-26", "2026-09-22", "sin_respuesta"),
    ("CORTAZAR",                "2026-08-26", "2026-09-22", "sin_respuesta"),
    ("SANOH",                   "2026-08-26", "2026-09-22", "sin_respuesta"),
    ("ACERLUM",                 "2026-08-26", "2026-09-22", "contesto_interes"),
    ("EXA INDUSTRIAL",          "2026-08-28", "2026-09-23", "sin_respuesta"),
    ("ESTROSA",                 "2026-08-28", "2026-09-23", "sin_respuesta"),
    ("DEMATIC",                 "2026-08-28", "2026-09-23", "sin_respuesta"),
    ("RC TOOLS",                "2026-08-28", None,         "contesto_no"),
    ("TEMPEL",                  "2026-08-28", None,         "reboto"),
    ("AMERTEC",                 "2026-08-28", "2026-09-23", "sin_respuesta"),
    ("SANHUA",                  "2026-08-28", None,         "reboto"),
    ("SELBOR",                  "2026-08-28", "2026-09-23", "sin_respuesta"),
    ("MANUFACTURAS TA",         "2026-08-28", None,         "sin_respuesta"),
    ("AEME",                    "2026-08-28", None,         "sin_respuesta"),
    ("INTERMEX",                "2026-09-01", None,         "reboto"),
    ("BOMBAS CENTRIFUGAS",      "2026-09-01", None,         "sin_respuesta"),
    ("CONTROL DIGITAL",         "2026-09-01", None,         "sin_respuesta"),
    ("PLOMAQ",                  "2026-09-01", None,         "sin_respuesta"),
    ("WABTEC",                  "2026-09-01", None,         "reboto"),
    ("SOIN3",                   "2026-09-01", None,         "sin_respuesta"),
    ("SERVICE TECH",            "2026-09-01", None,         "sin_respuesta"),
    ("GASODUCTOS",              "2026-09-01", None,         "sin_respuesta"),
    ("HIERRO Y EL ACERO",       "2026-09-01", None,         "sin_respuesta"),
    ("MATTSA",                  "2026-09-01", None,         "sin_respuesta"),
    ("GALVASID",                "2026-09-02", None,         "encaje_flojo"),
    ("LAMINA DESPLEGADA",       "2026-09-02", None,         "encaje_flojo"),
    ("GVS FILTER",              "2026-09-02", None,         "reboto"),
    ("COMBUSTION DIESEL",       "2026-09-02", None,         "sin_respuesta"),
    ("CONSTRUCCIONES METAL",    "2026-09-02", None,         "sin_respuesta"),
    ("PROYECTOS ELECTRICOS",    "2026-09-02", None,         "sin_respuesta"),
    ("GUMEX",                   "2026-09-02", None,         "sin_respuesta"),
    ("WELMON",                  "2026-09-02", None,         "sin_respuesta"),
    ("SUAREZ",                  "2026-09-02", None,         "reboto"),
    ("MAGNOACERO",              "2026-09-02", None,         "encaje_flojo"),
    ("SERVIACERO WORTHINGTON",  "2026-09-02", None,         "encaje_flojo"),
]


# El padron guarda los nombres en mayusculas y sin acentos. Estos son los que
# JP va a leer todos los dias en la pantalla, asi que van escritos como los
# escribiria una persona — y como se usaron en los correos.
NOMBRES = {
    "DUROSA": "Durosa",
    "AURIN": "Aurin de México",
    "CORTAZAR": "Maquinados de Precisión Cortazar",
    "SANOH": "Sanoh Industrial de México",
    "ACERLUM": "Acerlum Metalmecánica",
    "EXA INDUSTRIAL": "EXA Industrial",
    "ESTROSA": "Estrosa",
    "DEMATIC": "Dematic Logistics de México",
    "RC TOOLS": "RC Tools",
    "TEMPEL": "Tempel de México",
    "AMERTEC": "Amertec",
    "SANHUA": "Sanhua Automotive México",
    "SELBOR": "Industrias Selbor",
    "MANUFACTURAS TA": "Equipos y Manufacturas TA",
    "AEME": "Maquinados AEME",
    "INTERMEX": "Intermex Butt",
    "BOMBAS CENTRIFUGAS": "Bombas Centrífugas Alemanas",
    "CONTROL DIGITAL": "Control Digital",
    "PLOMAQ": "Plomaq",
    "WABTEC": "Wabtec",
    "SOIN3": "Soin3",
    "SERVICE TECH": "Service Tech de México",
    "GASODUCTOS": "Gasoductos y Estaciones del Norte",
    "HIERRO Y EL ACERO": "IMHASA",
    "MATTSA": "Mattsa Furnace Company",
    "GALVASID": "Galvasid",
    "LAMINA DESPLEGADA": "Lámina Desplegada",
    "GVS FILTER": "GVS Filter Technology de México",
    "COMBUSTION DIESEL": "Combustión Diesel de Monterrey",
    "CONSTRUCCIONES METAL": "Construcciones Metal Mecánicas Mexicanas",
    "PROYECTOS ELECTRICOS": "Proyectos Eléctricos y de Pailería CYG",
    "GUMEX": "Industrializaciones Gumex",
    "WELMON": "Welmon",
    "SUAREZ": "Empresas Suárez",
    "MAGNOACERO": "Grupo Industrial Magnoacero",
    "SERVIACERO WORTHINGTON": "Serviacero Worthington",
}

# El giro del DENUE traducido a los sectores que maneja la app.
def sector_de(scian: str, actividad: str) -> str:
    s = str(scian)
    if s.startswith(("3363", "3362", "3365", "336")):
        return "Vehículo comercial"
    if s.startswith(("3331", "3336")):
        return "Maquinaria pesada"
    if s.startswith("3323"):
        return "Construcción"
    if s.startswith("3334"):
        return "Energía / HVAC"
    if s.startswith(("3321", "3324", "3327", "3328", "3329", "333")):
        return "Industrial"
    return "Otro"


# Nombre y puesto de los contactos que sí conocemos.
PERSONAS = {
    "ACERLUM": [{"n": "Jesús Medina", "p": "Coordinador de Ingeniería",
                 "t": "", "e": "coordinador.ingenieria@acerlum.com"},
                {"n": "Adriana Torres", "p": "Coordinadora de Servicio a Cliente",
                 "t": "", "e": "adriana.torres@acerlum.com"}],
    "RC TOOLS": [{"n": "Feliciano", "p": "", "t": "", "e": "feliciano@rctools.com.mx"}],
}


def estado_y_accion(desenlace, seguimiento):
    """Estado de la cuenta, próxima acción y su fecha."""
    if desenlace == "contesto_interes":
        return ("Con interés",
                "Dar seguimiento: esperan definir pieza para temple o rectificado",
                "2026-09-30")
    if desenlace == "contesto_no":
        return ("Dormido", "", "")
    if desenlace == "reboto":
        return ("Sin contactar",
                "Conseguir correo correcto o llamar: el primer correo rebotó",
                "2026-09-25")
    if desenlace == "encaje_flojo":
        return ("Dormido", "", "")
    if seguimiento:
        return ("En acercamiento",
                "Llamar si no contesta el seguimiento de ruteo",
                "2026-09-30")
    return ("En acercamiento", "Enviar seguimiento de ruteo", "2026-09-24")


def notas_de(desenlace, envio, seguimiento, correo):
    lineas = [f"[{envio} · {VENDEDOR}] Correo de presentación enviado a {correo}."]
    if desenlace == "reboto":
        lineas.append(f"[{envio} · {VENDEDOR}] El correo REBOTÓ: la dirección o el "
                      f"dominio no existen. Hay que conseguir contacto por otra vía.")
    if seguimiento:
        lineas.append(f"[{seguimiento} · {VENDEDOR}] Seguimiento en el mismo hilo: "
                      f"se preguntó quién ve compras de maquinado o alta de proveedores.")
    if desenlace == "contesto_no":
        lineas.append(f"[2026-09-01 · {VENDEDOR}] Contestaron que no: se dedican al "
                      f"termoformado y tienen taller interno. El giro del DENUE estaba mal.")
    if desenlace == "contesto_interes":
        lineas.append(f"[2026-09-22 · {VENDEDOR}] Jesús Medina (ingeniería) ruteó con "
                      f"Adriana Torres, servicio a cliente.")
        lineas.append(f"[2026-09-22 · {VENDEDOR}] Adriana mandó su presentación: cortan, "
                      f"doblan, rolan, sueldan, maquinan y pintan. NO tienen temple ni "
                      f"rectificado. Clientes suyos: Progress Rail, IMESA, Pirelli, Metso, "
                      f"Gyptech.")
        lineas.append(f"[2026-09-23 · {VENDEDOR}] Se respondió ofreciendo participar en sus "
                      f"proyectos del norte con soldadura, ensamble y maquinado, más temple "
                      f"y rectificado. Brochure adjunto.")
    if desenlace == "encaje_flojo":
        lineas.append(f"[{envio} · {VENDEDOR}] Sin respuesta. Encaje flojo: su giro no "
                      f"compra maquinado. No se le da seguimiento.")
    return "\n".join(lineas)


def construir():
    c = pd.read_excel(XLSX, sheet_name="Campana").fillna("")
    registros = []
    for clave, envio, seguimiento, desenlace in CAMPANA:
        m = c[c["Nombre comercial"].str.upper().str.contains(clave, na=False)]
        if m.empty:
            print(f"  !! no encontré {clave} en el Excel", file=sys.stderr)
            continue
        r = m.iloc[0]
        correo = str(r["Correo"]).split(",")[0].strip()
        estado, accion, fecha_accion = estado_y_accion(desenlace, seguimiento)
        contactos = PERSONAS.get(clave) or [
            {"n": "", "p": "", "t": str(r["Teléfono"]).split(",")[0].strip(), "e": correo}]
        registros.append({
            "Empresa": NOMBRES.get(clave, str(r["Nombre comercial"]).title()),
            "Tipo": "Prospecto nuevo",
            "Sector": sector_de(r["SCIAN"], r["Actividad"]),
            "Vendedor": VENDEDOR,
            "Estado": estado,
            "Contactos": json.dumps(contactos, ensure_ascii=False),
            "ProximaAccion": accion,
            "FechaProximaAccion": fecha_accion,
            "FechaUltimoContacto": seguimiento or envio,
            "Notas": notas_de(desenlace, envio, seguimiento, correo),
        })
    return registros


def subir(rec):
    payload = json.dumps({"action": "add", "tab": "Cuentas", "rec": rec}).encode("utf-8")
    req = urllib.request.Request(
        API, data=payload, headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--subir", action="store_true")
    ap.add_argument("--solo", type=str, default=None, help="una sola empresa, para probar")
    args = ap.parse_args()

    regs = construir()
    if args.solo:
        regs = [r for r in regs if args.solo.upper() in r["Empresa"].upper()]

    print(f"{len(regs)} registros\n")
    for r in regs:
        print(f"  {r['Empresa'][:30]:31} {r['Estado'][:16]:17} {r['Sector'][:18]:19} "
              f"últ={r['FechaUltimoContacto']} prox={r['FechaProximaAccion'] or '—'}")

    if not args.subir:
        print("\n(simulación — nada se escribió; usa --subir para cargarlos)")
        sys.exit(0)

    print()
    for r in regs:
        try:
            resp = subir(r)
            print(f"  OK  {r['Empresa'][:34]:35} {resp.get('ok')}")
        except Exception as e:
            print(f"  ERROR {r['Empresa'][:32]:33} {e}")
