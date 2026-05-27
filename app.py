import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path
from datetime import datetime

st.set_page_config(page_title="Video Test Árbitros", page_icon="⚽", layout="wide")

PREGUNTAS_FILE = Path("preguntas.json")
RESULTADOS_DIR = Path("resultados")
RESULTADOS_DIR.mkdir(exist_ok=True)

RESUMEN_CSV = RESULTADOS_DIR / "resultados_resumen.csv"
DETALLE_CSV = RESULTADOS_DIR / "resultados_detalle.csv"
EXCEL_FILE = RESULTADOS_DIR / "resultados_videotest.xlsx"

DECISIONES = ["No falta", "Tiro libre directo", "Tiro libre indirecto", "Penal"]
SANCIONES = ["No tarjeta", "Amonestación", "Expulsión"]


def clave_admin():
    try:
        return st.secrets.get("ADMIN_PASSWORD", "admin123")
    except Exception:
        return "admin123"


def email_valido(email):
    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()) is not None


def cargar():
    with open(PREGUNTAS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def guardar(data):
    with open(PREGUNTAS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def faltan_respuestas(data):
    faltan = []
    for v in data["videos"]:
        if v.get("decision_correcta", "") not in DECISIONES:
            faltan.append(f'{v["titulo"]}: falta decisión técnica')
        if v.get("sancion_correcta", "") not in SANCIONES:
            faltan.append(f'{v["titulo"]}: falta sanción disciplinaria')
    return faltan


def nivel(porc):
    if porc >= 90:
        return "Excelente"
    if porc >= 80:
        return "Muy bueno"
    if porc >= 70:
        return "Bueno"
    if porc >= 60:
        return "Regular"
    return "Debe reforzar"


def leer_csv(path):
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


def guardar_seguro_csv(df, path, base_name):
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except PermissionError:
        nuevo = RESULTADOS_DIR / f"{base_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(nuevo, index=False, encoding="utf-8-sig")
        st.warning(f"{path.name} está abierto o bloqueado. Guardé copia: {nuevo.name}")


def guardar_seguro_excel(resumen, detalle):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            detalle.to_excel(writer, index=False, sheet_name="Detalle")
    except PermissionError:
        nuevo = RESULTADOS_DIR / f"resultados_videotest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(nuevo, engine="openpyxl") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            detalle.to_excel(writer, index=False, sheet_name="Detalle")
        st.warning(f"El Excel resultados_videotest.xlsx está abierto o bloqueado. Guardé copia: {nuevo.name}")


def corregir(data, respuestas):
    detalle = []
    puntos = 0
    total = 0

    for v in data["videos"]:
        pares = [
            ("Decisión técnica", respuestas[f'{v["id"]}_decision'], v["decision_correcta"], 2),
            ("Sanción disciplinaria", respuestas[f'{v["id"]}_sancion'], v["sancion_correcta"], 2),
        ]

        for pregunta, resp, correcta, pts in pares:
            ok = resp == correcta
            total += pts
            puntos += pts if ok else 0

            detalle.append({
                "tema": v.get("tema", ""),
                "subtema": v.get("subtema", ""),
                "video": v.get("titulo", ""),
                "pregunta": pregunta,
                "respuesta_usuario": resp,
                "respuesta_correcta": correcta,
                "correcta": "Sí" if ok else "No",
                "puntos": pts,
                "obtenido": pts if ok else 0,
                "criterio_admin": v.get("criterio_admin", ""),
                "explicacion_admin": v.get("explicacion_admin", ""),
            })

    porc = round((puntos / total) * 100, 2) if total else 0
    return puntos, total, porc, nivel(porc), detalle


def guardar_resultado(participante, puntos, total, porc, niv, detalle):
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    fila = {
        "fecha": fecha,
        **participante,
        "puntaje_obtenido": puntos,
        "puntaje_total": total,
        "porcentaje": porc,
        "nivel": niv,
    }

    resumen_nuevo = pd.DataFrame([fila])
    detalle_nuevo = pd.DataFrame(detalle)

    for k, v in fila.items():
        detalle_nuevo.insert(0, k, v)

    resumen = pd.concat([leer_csv(RESUMEN_CSV), resumen_nuevo], ignore_index=True)
    detalle_final = pd.concat([leer_csv(DETALLE_CSV), detalle_nuevo], ignore_index=True)

    guardar_seguro_csv(resumen, RESUMEN_CSV, "resultados_resumen")
    guardar_seguro_csv(detalle_final, DETALLE_CSV, "resultados_detalle")
    guardar_seguro_excel(resumen, detalle_final)


def admin():
    st.header("Panel administrador")

    pwd = st.text_input("Clave administrador", type="password")

    if pwd != clave_admin():
        st.info("Clave local inicial: admin123")
        return

    data = cargar()

    st.success("Acceso administrador correcto")
    st.info("Las respuestas correctas no salen de los videos automáticamente. Cargalas acá antes de tomar el test.")

    df = pd.DataFrame(data["videos"])

    df_edit = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        disabled=["id", "archivo"],
        column_config={
            "decision_correcta": st.column_config.SelectboxColumn(
                "decision_correcta",
                options=[""] + DECISIONES
            ),
            "sancion_correcta": st.column_config.SelectboxColumn(
                "sancion_correcta",
                options=[""] + SANCIONES
            ),
        },
    )

    if st.button("Guardar configuración de respuestas"):
        data["videos"] = df_edit.to_dict(orient="records")
        guardar(data)
        st.success("Configuración guardada")
        st.rerun()

    faltan = faltan_respuestas(data)

    if faltan:
        st.warning("Faltan respuestas correctas:")
        for x in faltan:
            st.write(f"- {x}")
    else:
        st.success("Banco completo. El test puede tomarse.")

    st.divider()

    if RESUMEN_CSV.exists():
        st.subheader("Resultados generales")
        st.dataframe(pd.read_csv(RESUMEN_CSV), use_container_width=True)

        if DETALLE_CSV.exists():
            st.subheader("Detalle de respuestas")
            st.dataframe(pd.read_csv(DETALLE_CSV), use_container_width=True)

        c1, c2, c3 = st.columns(3)

        c1.download_button(
            "Descargar resumen CSV",
            RESUMEN_CSV.read_bytes(),
            "resultados_resumen.csv"
        )

        if DETALLE_CSV.exists():
            c2.download_button(
                "Descargar detalle CSV",
                DETALLE_CSV.read_bytes(),
                "resultados_detalle.csv"
            )

        if EXCEL_FILE.exists():
            c3.download_button(
                "Descargar Excel",
                EXCEL_FILE.read_bytes(),
                "resultados_videotest.xlsx"
            )
    else:
        st.warning("Todavía no hay resultados guardados.")


def test():
    data = cargar()
    faltan = faltan_respuestas(data)

    if faltan:
        st.error("El test no está habilitado. El administrador debe cargar las respuestas correctas reales.")
        with st.expander("Ver faltantes"):
            for x in faltan:
                st.write(f"- {x}")
        return

    st.header("Ingreso del participante")

    with st.form("form_participante"):
        nombre = st.text_input("Nombre y apellido *")
        email = st.text_input("Correo electrónico *")
        categoria = st.selectbox(
            "Categoría / rol",
            ["Árbitro", "Árbitro asistente", "Cuarto árbitro", "Instructor", "Otro"]
        )
        institucion = st.text_input("Institución / departamento")
        aceptar = st.checkbox("Confirmo que realizaré el test de forma individual.")
        ok = st.form_submit_button("Comenzar test")

    if ok:
        if not nombre.strip():
            st.error("Debe ingresar nombre y apellido.")
            st.stop()

        if not email_valido(email):
            st.error("Debe ingresar un correo válido.")
            st.stop()

        if not aceptar:
            st.error("Debe confirmar que realizará el test de forma individual.")
            st.stop()

        st.session_state["datos_participante"] = {
            "nombre": nombre.strip(),
            "email": email.strip(),
            "categoria": categoria,
            "institucion": institucion.strip(),
        }
        st.session_state["test_habilitado"] = True
        st.rerun()

    if not st.session_state.get("test_habilitado", False):
        st.info("Completá tus datos y presioná Comenzar test.")
        return

    respuestas = {}
    st.header("Video test")

    with st.form("form_test"):
        for i, v in enumerate(data["videos"], start=1):
            st.subheader(f"Video {i}: {v['titulo']}")

            ruta = Path(v["archivo"])

            if ruta.exists():
                st.video(str(ruta))
            else:
                st.warning(f"No encontré el video: {v['archivo']}")

            respuestas[f'{v["id"]}_decision'] = st.radio(
                "1. Decisión técnica",
                DECISIONES,
                key=f'{v["id"]}_d',
                index=None
            ) or ""

            respuestas[f'{v["id"]}_sancion'] = st.radio(
                "2. Sanción disciplinaria",
                SANCIONES,
                key=f'{v["id"]}_s',
                index=None
            ) or ""

            st.divider()

        enviar = st.form_submit_button("Enviar respuestas y ver calificación")

    if enviar:
        if any(not r for r in respuestas.values()):
            st.error("Faltan preguntas por responder. Revisá el test antes de enviar.")
            st.stop()

        puntos, total, porc, niv, detalle = corregir(data, respuestas)
        guardar_resultado(st.session_state["datos_participante"], puntos, total, porc, niv, detalle)

        st.success("Test enviado correctamente")

        c1, c2, c3 = st.columns(3)
        c1.metric("Puntaje", f"{puntos} / {total}")
        c2.metric("Porcentaje", f"{porc}%")
        c3.metric("Calificación", niv)

        df_det = pd.DataFrame(detalle)

        resumen = df_det.groupby("tema", as_index=False).agg(
            puntos_obtenidos=("obtenido", "sum"),
            puntos_totales=("puntos", "sum")
        )
        resumen["porcentaje"] = (resumen["puntos_obtenidos"] / resumen["puntos_totales"] * 100).round(2)
        resumen["nivel"] = resumen["porcentaje"].apply(nivel)

        st.subheader("Resumen por tema")
        st.dataframe(resumen, use_container_width=True)

        st.info("Las respuestas correctas quedan solo para el administrador.")


st.title("⚽ Video Test Árbitros")
st.caption("Versión corregida: formulario de participante sin conflicto de session_state.")

menu = st.sidebar.radio("Menú", ["Realizar test", "Administrador"])

if menu == "Administrador":
    admin()
else:
    test()
