"""Registra en el tablero lo que se hizo el 23-sep-2026.

Cinco correos de seguimiento enviados y el alta de ACV Group. La bitácora se
APENDE, nunca se reemplaza: se lee lo que ya había y se le agrega la línea nueva.
"""
import json
import urllib.request

API = ("https://script.google.com/macros/s/"
       "AKfycbym1-YyG7LPm2fOxSrad50vxDHHiHbKRDaJLJ5jVF-T83N_kbqN9SZJuL75pwNJ2slCYw/exec")
V = "Juan Pablo Lopez"
HOY = "2026-09-23"
SEM = "2026-09-30"   # el plazo de una semana que pidió JP


def leer():
    with urllib.request.urlopen(API + "?action=all&t=99", timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def post(payload):
    req = urllib.request.Request(
        API, data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read().decode("utf-8"))


def nota(previo, texto):
    linea = f"[{HOY} · {V}] {texto}"
    return (str(previo).rstrip() + "\n" + linea) if str(previo).strip() else linea


# Lo que cambia en cada cuenta que se tocó hoy.
CUENTAS = {
    "Madero System": dict(
        texto="Se retomó el hilo de las muestras M6. Quedaron aceptadas; se pidió "
              "la propuesta de ajustes de diseño para bajar costo, que el cliente "
              "mencionó y nunca envió. Si no hay respuesta en una semana, se cierra "
              "la oportunidad como perdida.",
        accion="Sin respuesta al 30-sep: cerrar la oportunidad como perdida",
        fecha=SEM),
    "Wittur Mx": dict(
        texto="Al revisar el hilo se encontró que la pelota estaba de NUESTRO lado: "
              "el 13-ago Fátima pidió cotizar astragales de aluminio y mandó dibujos, "
              "y la cotización nunca salió. Se reconoció la demora y se ofreció "
              "trabajarlos si el requerimiento sigue vigente.",
        accion="Cotizar astragales de aluminio en cuanto confirmen tolerancias",
        fecha=SEM, estado="Con interés"),
    "Kimball": dict(
        texto="Cuarto correo, esta vez de cierre: se le dio salida para decir que no. "
              "Van tres correos previos sin respuesta (8, 12 y 19 de junio).",
        accion="Llamar al 871 750 9818. Si no contesta, archivar",
        fecha=SEM, estado="En acercamiento"),
    "Marelli": dict(
        texto="Se escribió a diego.gomez@marelli.com preguntando quién ve maquinado. "
              "OJO: el correo que traía el tablero estaba mal; el bueno es "
              "diego.gomez@, no diego.cortes@.",
        accion="Esperar a que rutee con el área de maquinado",
        fecha=SEM, estado="En acercamiento",
        contactos=[{"n": "Diego Gómez", "p": "Purchasing", "t": "",
                    "e": "diego.gomez@marelli.com"}]),
    "Fox Steel": dict(
        texto="Se pidió a info@foxsteel.mx el correo correcto de compras. Los dos "
              "intentos de escribirle a Lilia García rebotaron: lgarcia@ el 8-jun y "
              "lilia.garcia@ el 24-jun. Nunca le ha llegado nada.",
        accion="Esperar el correo de compras que nos pasen",
        fecha=SEM, estado="En acercamiento"),
}

ACV = {
    "Empresa": "ACV Group",
    "Tipo": "Prospecto nuevo",
    "Sector": "Industrial",
    "Vendedor": V,
    "Estado": "Con interés",
    "Contactos": json.dumps(
        [{"n": "Enrique Pérez", "p": "", "t": "", "e": "enrique.perez@acvgroup.com.mx"}],
        ensure_ascii=False),
    "ProximaAccion": "Pedirles el dibujo de una pieza representativa para cotizar",
    "FechaProximaAccion": SEM,
    "FechaUltimoContacto": "2026-09-14",
    "Notas": (
        f"[2026-09-14 · {V}] Llegó solicitud formal de información de Enrique Pérez: "
        f"están ampliando líneas y buscan proveedores de maquila — corte láser, "
        f"plasma, doblez, rolado, maquinado y operaciones complementarias. Pidieron "
        f"capacidades instaladas, equipos, dimensiones y espesores máximos, "
        f"materiales, tolerancias y tiempos de respuesta.\n"
        f"[2026-09-14 · {V}] Se firmó el NDA con ACV Group.\n"
        f"[{HOY} · {V}] Alta en el tablero. De los seis procesos que piden, sólo "
        f"maquinado es propio; láser, plasma y doblez van con tercero. El gancho "
        f"fuerte es lo que NO pidieron: temple por inducción y rectificado."),
}


if __name__ == "__main__":
    d = leer()
    por_nombre = {c["Empresa"]: c for c in d["cuentas"]}

    for empresa, cambio in CUENTAS.items():
        c = por_nombre.get(empresa)
        if not c:
            print(f"  !! no encontré la cuenta {empresa}")
            continue
        rec = {
            "Notas": nota(c.get("Notas", ""), cambio["texto"]),
            "ProximaAccion": cambio["accion"],
            "FechaProximaAccion": cambio["fecha"],
            "FechaUltimoContacto": HOY,
        }
        if "estado" in cambio:
            rec["Estado"] = cambio["estado"]
        if "contactos" in cambio:
            rec["Contactos"] = json.dumps(cambio["contactos"], ensure_ascii=False)
        r = post({"action": "update", "tab": "Cuentas", "id": c["ID"], "rec": rec})
        print(f"  cuenta  {empresa:16} -> {r.get('ok')}")

    # La oportunidad de Madero lleva su propio plazo.
    op = next((o for o in d["oportunidades"]
               if o["Empresa"] == "Madero System" and o["Etapa"] == "Cotización enviada"), None)
    if op:
        rec = {
            "Notas": nota(op.get("Notas", ""),
                          "Se pidió la propuesta de ajustes de diseño. Plazo al 30-sep; "
                          "sin respuesta, se cierra como perdida."),
            "ProximaAccion": "Cerrar como perdida si no llegó la propuesta de diseño",
            "FechaProximaAccion": SEM,
            "FechaUltimoContacto": HOY,
        }
        r = post({"action": "update", "tab": "Oportunidades", "id": op["ID"], "rec": rec})
        print(f"  oportunidad Madero System -> {r.get('ok')}")

    if "ACV Group" in por_nombre:
        print("  ACV Group ya existía, no se duplica")
    else:
        r = post({"action": "add", "tab": "Cuentas", "rec": ACV})
        print(f"  alta    ACV Group        -> {r.get('ok')}")
